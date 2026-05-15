const fileInput = document.getElementById("fileInput");
const tableContainer = document.getElementById("tableContainer");
const sheetButtons = document.getElementById("sheetButtons");

let workbook;

fileInput.addEventListener("change", handleFileUpload);

function handleFileUpload(event) {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

  const reader = new FileReader();

  reader.onload = function (event) {
    const data = event.target.result;

    workbook = XLSX.read(data, {
      type: "array"
    });

    createSheetButtons();
    displaySheet(workbook.SheetNames[0]);
  };

  reader.readAsArrayBuffer(file);
}

function createSheetButtons() {
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