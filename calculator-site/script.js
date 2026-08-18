const display = document.querySelector("#display");
const keys = document.querySelectorAll(".key");

let expression = "";
let justEvaluated = false;

const operators = new Set(["+", "-", "*", "/", "%"]);

function render(value) {
  const text = value || "0";
  display.textContent = text.length > 15 ? text.slice(0, 15) : text;
}

function appendValue(value) {
  if (justEvaluated && !operators.has(value)) {
    expression = "";
  }

  justEvaluated = false;
  const last = expression.at(-1);

  if (value === "." && currentNumber().includes(".")) {
    return;
  }

  if (operators.has(value)) {
    if (!expression && value !== "-") {
      return;
    }

    if (operators.has(last)) {
      expression = expression.slice(0, -1);
    }
  }

  expression += value;
  render(formatForDisplay(expression));
}

function currentNumber() {
  return expression.split(/[+\-*/%]/).at(-1) || "";
}

function formatForDisplay(value) {
  return value.replaceAll("*", "x");
}

function clearAll() {
  expression = "";
  justEvaluated = false;
  render(expression);
}

function deleteLast() {
  expression = expression.slice(0, -1);
  justEvaluated = false;
  render(formatForDisplay(expression));
}

function calculate() {
  if (!expression || operators.has(expression.at(-1))) {
    return;
  }

  try {
    const result = Function(`"use strict"; return (${expression})`)();

    if (!Number.isFinite(result)) {
      throw new Error("Invalid result");
    }

    expression = Number.isInteger(result)
      ? String(result)
      : String(Number(result.toFixed(10)));
    justEvaluated = true;
    render(expression);
  } catch {
    expression = "";
    justEvaluated = true;
    render("Error");
  }
}

function pressButton(button) {
  if (!button) {
    return;
  }

  button.classList.add("is-active");
  window.setTimeout(() => button.classList.remove("is-active"), 120);
}

keys.forEach((key) => {
  key.addEventListener("click", () => {
    const action = key.dataset.action;
    const value = key.dataset.value;

    pressButton(key);

    if (action === "clear") {
      clearAll();
    } else if (action === "delete") {
      deleteLast();
    } else if (action === "equals") {
      calculate();
    } else if (value) {
      appendValue(value);
    }
  });
});

window.addEventListener("keydown", (event) => {
  const { key } = event;
  const lookup = {
    Enter: "[data-action='equals']",
    "=": "[data-action='equals']",
    Backspace: "[data-action='delete']",
    Escape: "[data-action='clear']",
    c: "[data-action='clear']",
    C: "[data-action='clear']",
  };

  if (/^[0-9.+\-*/%]$/.test(key)) {
    event.preventDefault();
    appendValue(key);
    pressButton(document.querySelector(`[data-value="${CSS.escape(key)}"]`));
    return;
  }

  const selector = lookup[key];

  if (selector) {
    event.preventDefault();
    const button = document.querySelector(selector);
    pressButton(button);
    button.click();
  }
});

render(expression);
