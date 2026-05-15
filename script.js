const fileInput = document.getElementById("fileInput");
const tableContainer = document.getElementById("tableContainer");

fileInput.addEventListener("change", handleFileUpload);

function handleFileUpload(event) {
  const file = event.target.files[0];

  if (!file) {
    return;
  }

  const reader = new FileReader();

  reader.onload = function (event) {
    const data = event.target.result;

    const workbook = XLSX.read(data, {
      type: "array"
    });

    const firstSheetName = workbook.SheetNames[0];
    const firstSheet = workbook.Sheets[firstSheetName];

    const htmlTable = XLSX.utils.sheet_to_html(firstSheet);

    tableContainer.innerHTML = htmlTable;
  };

  reader.readAsArrayBuffer(file);
}