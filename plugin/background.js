chrome.action.onClicked.addListener(async (tab) => {
  try {
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: await copyQuestionAndAnswer,
    });
  } catch (err) {
    console.error("Failed to execute script:", err);
  }
});

async function copyQuestionAndAnswer() {
  // Get question
  const questionElement = document.querySelector("div.ck-content p");
  let questionText = questionElement ? questionElement.textContent.trim() : "";

  // Get answer fieldset
  const fieldsets = document.getElementsByTagName("fieldset");
  let answerFieldset = null;

  for (let fieldset of fieldsets) {
    const legend = fieldset.querySelector("legend");
    if (legend && legend.textContent === "Antwort") {
      answerFieldset = fieldset;
      break;
    }
  }

  let formattedText = "";

  if (answerFieldset) {
    // Check if it's a table-based matching question
    const table = answerFieldset.querySelector("table");

    if (table) {
      // Handle matching table type question
      const rows = table.querySelectorAll("tbody tr");
      let terms = [];
      let answers = [];

      // Skip the header row (index 0)
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const term = row.querySelector(".fixcol p strong")?.textContent.trim();
        const answer = row
          .querySelector(".sourcecol p span")
          ?.textContent.trim();

        if (term && answer) {
          terms.push(term);
          answers.push(answer);
        }
      }

      // Format in the desired structure using numbers and letters
      const formattedTerms = terms
        .map((term, index) => `${index + 1}. "${term}"`)
        .join("\n");

      const formattedAnswers = answers
        .map(
          (answer, index) => `${String.fromCharCode(97 + index)}. "${answer}"`
        )
        .join("\n");

      formattedText = `${formattedTerms}\n\n${formattedAnswers}`;
    } else {
      // Handle original gap text type question
      const selects = answerFieldset.querySelectorAll("select");
      let answerText = answerFieldset.querySelector("p").textContent;

      // Process each select element
      selects.forEach((select) => {
        const options = Array.from(select.options)
          .filter((option) => option.text.trim() !== "")
          .map((option) => option.text.trim());

        // Replace the select content with the options in brackets
        answerText = answerText.replace(
          select.textContent,
          `[${options.join(" | ")}]`
        );
      });

      formattedText = `${questionText}\n${answerText.trim()}`;
    }

    console.log({ formattedText });

    // Send the formatted text to your local server
    await fetch("http://localhost:5000/log", {
      method: "POST",
      mode: "cors",
      credentials: "omit",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Origin: chrome.runtime.getURL(""),
      },
      body: JSON.stringify({
        question: questionText,
        answer: formattedText,
        formatted: formattedText,
        type: table ? "matching" : "gaptext",
      }),
    });
  }
}
