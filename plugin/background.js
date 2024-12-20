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
  const questionElement = document.querySelector("form.fixform");
  const questionContent = questionElement.querySelector(".ck-content");
  let questionText = questionContent ? questionContent.textContent.trim() : "";

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
      // Handle table-based matching questions
      const rows = table.querySelectorAll("tbody tr");
      let terms = [];
      let answers = [];

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const term = row.querySelector(".fixcol p")?.textContent.trim();
        const answer = row.querySelector(".sourcecol p")?.textContent.trim();

        if (term && answer) {
          terms.push(term);
          answers.push(answer);
        }
      }

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
      // Handle gap text questions
      const contentDiv = answerFieldset.querySelector(".ck-content");
      const textContainer = contentDiv.querySelector("p, pre");
      let answerText = "";

      if (textContainer) {
        // Create a document fragment to work with
        const container = document.createElement("div");
        container.innerHTML = textContainer.innerHTML;

        // Process each select element
        const selects = container.querySelectorAll("select");
        selects.forEach((select) => {
          const options = Array.from(select.options)
            .filter((option) => option.text.trim() !== "")
            .map((option) => option.text.trim());

          // Create a placeholder for the select element
          const placeholder = document.createTextNode(
            `[${options.join(" | ")}]`
          );
          select.parentNode.replaceChild(placeholder, select);
        });

        // Get the text content with replaced selects
        answerText = container.textContent.trim();
      }

      formattedText = `${questionText}\n${answerText}`;
    }

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
