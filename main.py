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
from flask import Flask, request, jsonify
import signal

if sys.version_info < (3, 8):
    import importlib_metadata
else:
    import importlib.metadata as importlib_metadata


class Paths:
    def __init__(self, path) -> Path:
        self.path = Path(path).absolute()

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

    @property
    def changed(self) -> bool:
        if self.stopped:
            return False

        is_changed = False
        for path, modified in self.modified.items():
            self.modified[path] = os.path.getmtime(path)
            if not is_changed and self.modified[path] != modified:
                is_changed = True
        return is_changed

    def update_modified(self):
        for path in self.watch_paths:
            self.modified[path] = os.path.getmtime(path)

    def stop(self):
        self.stopped = True

    def start(self):
        self.update_modified()
        self.stopped = False


class Shared:
    paths: Paths
    docx_watcher: FilesWatcher
    xmls_watcher: FilesWatcher
    cmds: queue.Queue


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


def process_with_claude(question, type):
    try:
        client = anthropic.Anthropic(
            api_key=""
        )

        if type == "gaptext":
            prompt = f"""
                   For each [...] bracket, return only ONE correct answer.
                   Return answers separated by comma.
                   Example input: [big | small] elephant drinks [hot | cold | warm] water
                   Example output: small, cold
                   Do not include any explanations, the terms, or additional text.

                   {question}
                   """
        elif type == "matching":
            prompt = f"""
                        Please analyze this matching question and match the numbered items with their corresponding letters. Return only the number -> letter pairs in the format:
                        Example output:
                        1 -> b
                        2 -> c
                        3 -> a
                        ...
                        Do not include any explanations, the terms, or additional text.

                        {question}
                        """

        message = client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=1000,
            temperature=0,
            messages=[
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )

        return message.content[0].text

    except Exception as e:
        return str(e)


def run_flask_app():
    app = Flask(__name__)

    @app.after_request
    def after_request(response):
        response.headers.add('Access-Control-Allow-Origin', '*')
        response.headers.add('Access-Control-Allow-Headers', 'Content-Type,Authorization')
        response.headers.add('Access-Control-Allow-Methods', 'POST,OPTIONS')
        return response

    @app.route('/log', methods=['POST', 'OPTIONS'])
    def log_request():
        if request.method == 'OPTIONS':
            return jsonify({"status": "ok"}), 200

        try:
            content = request.get_json(silent=True)
            print(f"Question: {content.get('question', 'N/A')}")
            print(f"Answer: {content.get('answer', 'N/A')}")
            print(f"Type: {content.get('type', 'N/A')}")
            response = process_with_claude(content.get('answer', 'N/A'), content.get('type', 'N/A'))
            print(f"Response: {response}")

            document_xml_path = Shared.paths.ext_xmls[0]

            modify_xml_file(document_xml_path, response)

            # Trigger an update in the Word document
            # Shared.cmds.put("update")

            return jsonify({"status": "success", "message": "Request logged and document updated"}), 200

        except Exception as e:
            print(f"Error logging request: {str(e)}")
            return jsonify({"status": "error", "message": str(e)}), 500

    app.run(host='localhost', port=5000, debug=False)


def watcher_thread(cmds: queue.Queue):
    while True:
        if Shared.docx_watcher.changed:
            print("Change in docx file detected!")
            cmds.put("reload")

        if Shared.xmls_watcher.changed:
            print("Change in extracted xmls detected!")
            cmds.put("update")

        time.sleep(0.1)


def update(parser):
    try:
        for path in Shared.paths.ext_xmls:
            etree.parse(path, parser)
    except etree.XMLSyntaxError:
        pass

    print("updating..")

    with zipfile.ZipFile(Shared.paths.path, "w") as file:
        for path in Shared.paths.ext_dirpath.glob("**/*"):
            file.write(path, path.relative_to(Shared.paths.ext_dirpath))


def preview_thread(cmds: queue.Queue):
    word_app = win32.Dispatch("Word.Application", pythoncom.CoInitialize())
    word_app.Visible = True

    parser = etree.XMLParser(remove_blank_text=True)
    doc = run_preview(word_app, None, parser)

    Shared.docx_watcher.start()
    Shared.xmls_watcher.start()

    while True:
        try:
            cmd = cmds.get(timeout=1)
        except queue.Empty:
            cmd = None

            try:
                word_app.Visible
            except pythoncom.com_error:
                print()
                print("Word was closed. exiting..")
                exit()

        if cmd == "reload":
            print("reloading..")
            doc = run_preview(word_app, doc, parser)

        elif cmd == "update":
            update(parser)

        elif cmd == "quit":
            print("exiting..")
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
        description="Preview a Docx file in MS Word with live updates from Flask API.",
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

    # Start Flask app in a separate thread
    threading.Thread(target=run_flask_app, daemon=True).start()


if __name__ == "__main__":
    main()
