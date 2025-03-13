import queue
import shutil
import threading
import argparse
import zipfile
import win32com.client as win32
import os
import sys
import time
from pathlib import Path
import pythoncom
from lxml import etree
import anthropic
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

if sys.version_info < (3, 8):
    import importlib_metadata
else:
    import importlib.metadata as importlib_metadata


class Paths:
    def __init__(self, path) -> Path:
        self.path = Path(path).absolute()
        self.question_path = self.path.parent / "question.txt"

    @property
    def preview_copy_path(self) -> Path:
        return self.path.with_stem(self.path.stem + "__preview")

    @property
    def ext_dirpath(self) -> Path:
        return self.path.with_name(self.path.name + "__extracted")

    @property
    def ext_xmls(self) -> list[Path]:
        return [
            self.ext_dirpath / "word" / "document.xml",
            self.ext_dirpath / "word" / "styles.xml",
        ]


class FilesWatcher:
    def __init__(self, watch_paths: list[Path]):
        self.watch_paths = watch_paths
        self.modified = {}
        self.stopped = True
        self.ignore_next_change = False

    @property
    def changed(self) -> bool:
        if self.stopped or self.ignore_next_change:
            if self.ignore_next_change:
                # Reset flag after ignoring one change
                self.ignore_next_change = False
            return False

        is_changed = False
        for path, modified in self.modified.items():
            current_mod_time = os.path.getmtime(path)
            if not is_changed and current_mod_time != modified:
                is_changed = True
                self.modified[path] = current_mod_time
        return is_changed

    def update_modified(self):
        for path in self.watch_paths:
            self.modified[path] = os.path.getmtime(path)

    def stop(self):
        self.stopped = True

    def start(self):
        self.update_modified()
        self.stopped = False


class QuestionFileHandler(FileSystemEventHandler):
    def __init__(self):
        super().__init__()
        self.last_processed_time = 0
        self.debounce_seconds = 1.0  # Debounce time to avoid multiple rapid triggers

    def on_modified(self, event):
        current_time = time.time()
        # Only process if enough time has passed since last processing
        if (current_time - self.last_processed_time) > self.debounce_seconds:
            if event.src_path == str(Shared.paths.question_path):
                print(f"Question file modified: {event.src_path}")
                with open(event.src_path, "r", encoding="utf-8") as file:
                    content = file.read().strip()
                    process_question(content)
                self.last_processed_time = current_time


class Shared:
    paths: Paths
    docx_watcher: FilesWatcher
    xmls_watcher: FilesWatcher
    cmds: queue.Queue
    update_lock = threading.Lock()
    operation_in_progress = False


def modify_xml_file(filepath, text):
    with open(filepath, 'r', encoding='utf-8') as file:
        content = file.read()

    start_tag = '<w:t answer="true">'

    # Check if the tag exists
    if start_tag not in content:
        # Tag doesn't exist, use the modify_xml_file logic
        with open(filepath, 'r', encoding='utf-8') as file:
            content_lines = file.readlines()

        target_line = -1
        for i, line in enumerate(content_lines):
            if '<w:sectPr' in line:
                target_line = i
                break

        if target_line == -1:
            raise Exception("Target tag <w:sectPr> not found in the file")

        content_lines.insert(target_line, f'<w:p><w:r><w:t answer="true">{text}</w:t></w:r></w:p>\n')

        with open(filepath, 'w', encoding='utf-8') as file:
            file.writelines(content_lines)

    else:
        # Tag exists, use the replace_xml_content logic
        end_tag = '</w:t>'

        start_pos = content.find(start_tag)
        content_start = start_pos + len(start_tag)

        end_pos = content.find(end_tag, content_start)
        if end_pos == -1:
            raise Exception("End tag not found")

        new_file_content = (
                content[:content_start] +
                text +
                content[end_pos:]
        )

        with open(filepath, 'w', encoding='utf-8') as file:
            file.write(new_file_content)

    return True


def process_with_claude(question):
    try:
        client = anthropic.Anthropic(
            api_key=""
        )

        message = client.messages.create(
            model="claude-3-7-sonnet-20250219",
            max_tokens=10000,
            temperature=0,
            messages=[
                {
                    "role": "user",
                    "content": question
                }
            ]
        )

        print(message)

        return message.content[0].text

    except Exception as e:
        return str(e)


def process_question(question):
    # Use a lock to prevent multiple operations happening at once
    with Shared.update_lock:
        if Shared.operation_in_progress:
            print("Operation already in progress, skipping...")
            return

        Shared.operation_in_progress = True

        try:
            print(f"Processing question: {question}")
            response = process_with_claude(question)
            print(f"Claude's response: {response}")

            # Tell watchers to ignore changes we're about to make
            Shared.docx_watcher.ignore_next_change = True
            Shared.xmls_watcher.ignore_next_change = True

            document_xml_path = Shared.paths.ext_xmls[0]
            modify_xml_file(document_xml_path, response)

            # Use a single update command to avoid multiple reload cycles
            Shared.cmds.put("update_single")
        finally:
            Shared.operation_in_progress = False


def watcher_thread(cmds: queue.Queue):
    # Add debounce mechanism
    last_docx_change = 0
    last_xml_change = 0
    debounce_time = 1.0  # seconds

    while True:
        current_time = time.time()

        if Shared.docx_watcher.changed:
            # Only act on changes that are not too close together
            if current_time - last_docx_change > debounce_time:
                print("Change in docx file detected!")
                cmds.put("reload")
                last_docx_change = current_time

        if Shared.xmls_watcher.changed:
            # Only act on changes that are not too close together
            if current_time - last_xml_change > debounce_time:
                print("Change in extracted xmls detected!")
                cmds.put("update")
                last_xml_change = current_time

        time.sleep(0.1)


def update(parser, prevent_reload=False):
    try:
        for path in Shared.paths.ext_xmls:
            etree.parse(path, parser)
    except etree.XMLSyntaxError:
        pass

    print("updating..")

    # If we want to prevent the update from triggering a reload
    if prevent_reload:
        Shared.docx_watcher.ignore_next_change = True

    with zipfile.ZipFile(Shared.paths.path, "w") as file:
        for path in Shared.paths.ext_dirpath.glob("**/*"):
            file.write(path, path.relative_to(Shared.paths.ext_dirpath))


def setup_question_file_watcher():
    # Create the question.txt file if it doesn't exist
    if not Shared.paths.question_path.exists():
        with open(Shared.paths.question_path, 'w', encoding='utf-8') as f:
            f.write("Enter your question here and save the file.")

    # Set up the file watcher
    event_handler = QuestionFileHandler()
    observer = Observer()
    observer.schedule(event_handler, str(Shared.paths.question_path.parent), recursive=False)
    observer.start()
    print(f"Watching question file at: {Shared.paths.question_path}")

    return observer


def preview_thread(cmds: queue.Queue):
    word_app = win32.Dispatch("Word.Application", pythoncom.CoInitialize())
    word_app.Visible = True

    parser = etree.XMLParser(remove_blank_text=True)
    doc = run_preview(word_app, None, parser)

    Shared.docx_watcher.start()
    Shared.xmls_watcher.start()

    # Set up the question file watcher
    observer = setup_question_file_watcher()

    # Keep track of when last reload happened to prevent rapid reloads
    last_reload_time = 0
    reload_debounce_time = 2.0  # seconds

    while True:
        current_time = time.time()

        try:
            cmd = cmds.get(timeout=1)
        except queue.Empty:
            cmd = None

            try:
                word_app.Visible
            except pythoncom.com_error:
                print()
                print("Word was closed. exiting..")
                observer.stop()
                observer.join()
                exit()

        if cmd == "reload" and (current_time - last_reload_time > reload_debounce_time):
            print("reloading..")
            last_reload_time = current_time
            doc = run_preview(word_app, doc, parser)

        elif cmd == "update":
            update(parser)

        elif cmd == "update_single":
            # Special command for single updates that shouldn't trigger reload cascade
            update(parser, prevent_reload=True)

        elif cmd == "quit":
            print("exiting..")
            observer.stop()
            observer.join()
            try:
                if doc in word_app.Documents:
                    doc.Close()
                word_app.Quit()
            except pythoncom.com_error:
                pass
            break


def run_preview(word_app, doc, parser):
    if doc is not None and doc in word_app.Documents:
        doc.Close()

    shutil.copyfile(Shared.paths.path, Shared.paths.preview_copy_path)
    doc = word_app.Documents.Open(str(Shared.paths.preview_copy_path))

    Shared.xmls_watcher.stop()
    with zipfile.ZipFile(Shared.paths.path) as file:
        for name in file.namelist():
            file.extract(name, Shared.paths.ext_dirpath)
        for path in Shared.paths.ext_xmls:
            tree = etree.parse(path, parser)
            tree.write(path, pretty_print=True, encoding="utf-8")
    Shared.xmls_watcher.start()

    return doc


def input_thread(cmds: queue.Queue):
    try:
        print("Press 'r' to reload. 'q' to quit.")
        while True:
            cmd = input("> ")
            if cmd == "q":
                cmds.put("quit")
                break
            elif cmd == "r":
                cmds.put("reload")
    except KeyboardInterrupt:
        cmds.put("quit")
        pass


def main():
    if os.name != "nt":
        print("Currently, this tool is Windows only.")
        return

    parser = argparse.ArgumentParser(
        prog="docx-live-reload",
        description="Preview a Docx file in MS Word with live updates from question.txt.",
        epilog="Created by idtareq@gmail.com",
    )

    def check_file(filename):
        if not Path(filename).suffix == ".docx":
            parser.error("File must be of the Docx format")
        elif not Path(filename).exists():
            parser.error("File does not exist")
        else:
            return filename

    parser.add_argument(
        "-V",
        "--version",
        action="version",
        version=f"%(prog)s {importlib_metadata.version('docx-live-reload')}",
    )
    parser.add_argument("docx_path", type=check_file)

    args = parser.parse_args()

    Shared.paths = Paths(args.docx_path)
    Shared.docx_watcher = FilesWatcher([Shared.paths.path])
    Shared.xmls_watcher = FilesWatcher(Shared.paths.ext_xmls)
    Shared.cmds = queue.Queue()

    # Start all threads
    threading.Thread(target=watcher_thread, daemon=True, args=(Shared.cmds,)).start()
    threading.Thread(target=input_thread, daemon=True, args=(Shared.cmds,)).start()
    threading.Thread(target=preview_thread, args=(Shared.cmds,)).start()


if __name__ == "__main__":
    main()