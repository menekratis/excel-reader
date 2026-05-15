const fileInput = document.getElementById("fileInput");
const tableContainer = document.getElementById("tableContainer");
const sheetButtons = document.getElementById("sheetButtons");

let workbook = null;

window.addEventListener("DOMContentLoaded", function () {
  loadDefaultWorkbook();
});

fileInput.addEventListener("change", function (event) {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

  readUploadedFile(file);
});

function loadDefaultWorkbook() {
  if (typeof DEFAULT_WORKBOOK_DATA === "undefined") {
    showMessage("Default data file was not found.");
    return;
  }

  if (!DEFAULT_WORKBOOK_DATA || !DEFAULT_WORKBOOK_DATA.sheets) {
    showMessage("Default data is empty or invalid.");
    return;
  }

  workbook = XLSX.utils.book_new();

  DEFAULT_WORKBOOK_DATA.sheets.forEach(function (sheetData) {
    const worksheet = XLSX.utils.aoa_to_sheet(sheetData.rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetData.name);
  });

  displaySheetButtons();
  displaySheet(workbook.SheetNames[0]);
}

function readUploadedFile(file) {
  const reader = new FileReader();

  reader.onload = function (event) {
    const data = event.target.result;

    workbook = XLSX.read(data, {
      type: "array"
    });

    displaySheetButtons();
    displaySheet(workbook.SheetNames[0]);
  };

  reader.readAsArrayBuffer(file);
}

function displaySheetButtons() {
  sheetButtons.innerHTML = "";

  workbook.SheetNames.forEach(function (sheetName) {
    const button = document.createElement("button");
    button.textContent = sheetName;

    button.addEventListener("click", function () {
      displaySheet(sheetName);
    });

    sheetButtons.appendChild(button);
  });
}

function displaySheet(sheetName) {
  const sheet = workbook.Sheets[sheetName];
  const htmlTable = XLSX.utils.sheet_to_html(sheet);

  tableContainer.innerHTML = htmlTable;
  updateActiveButton(sheetName);
}

function updateActiveButton(activeSheetName) {
  const buttons = sheetButtons.querySelectorAll("button");

  buttons.forEach(function (button) {
    if (button.textContent === activeSheetName) {
      button.classList.add("active");
    } else {
      button.classList.remove("active");
    }
  });
}

function showMessage(message) {
  tableContainer.innerHTML = `<p class="empty-message">${message}</p>`;
}