const xlsx = require("xlsx");
const path = require("path");

// Nazwa pliku wejściowego
const inputFile = path.resolve("material_reuse.xlsx");

// Wyciągnij nazwę pliku bez rozszerzenia
const fileName = path.basename(inputFile, ".xlsx");

// Nazwa pliku wyjściowego
const outputFile = path.resolve(`sorted_${fileName}.xlsx`);

// Wczytaj plik
const workbook = xlsx.readFile(inputFile);
const sheet = workbook.Sheets[workbook.SheetNames[0]];

// Zamień dane na tablicę obiektów
const data = xlsx.utils.sheet_to_json(sheet);

// Sortuj alfabetycznie po polu "categories"
const sortedData = [...data].sort((a, b) => {
  const catA = a["categories"]?.toString().toLowerCase() || "";
  const catB = b["categories"]?.toString().toLowerCase() || "";
  return catA.localeCompare(catB);
});

console.log(sortedData)

// Tworzenie nowego workbooka i arkusza
const newWorkbook = xlsx.utils.book_new();
const newSheet = xlsx.utils.json_to_sheet(sortedData);
xlsx.utils.book_append_sheet(newWorkbook, newSheet, "Sorted");

// Zapisz nowy plik
xlsx.writeFile(newWorkbook, outputFile);
console.log(`Zapisano jako: ${outputFile}`);
