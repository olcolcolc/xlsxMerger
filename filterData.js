const xlsx = require("xlsx");
const path = require("path");

// Ścieżka do wejściowego i wyjściowego pliku
const inputFile = path.resolve("material_reuse.xlsx"); // zmień nazwę
const outputFile = path.resolve("sorted.xlsx");

// Wczytaj plik
const workbook = xlsx.readFile(inputFile);
const sheet = workbook.Sheets[workbook.SheetNames[0]];

// Zamień dane na tablicę obiektów
const data = xlsx.utils.sheet_to_json(sheet); // header: 0 domyślnie

// Grupowanie po polu "categoria"
const sortedList = {};

data.forEach((item) => {
  const category = item["categoria"];
  if (!sortedList[category]) {
    sortedList[category] = [];
  }
  sortedList[category].push(item);
});

// Tworzenie nowego pliku Excel z osobnymi arkuszami
const newWorkbook = xlsx.utils.book_new();

Object.entries(sortedList).forEach(([category, items]) => {
  const sheet = xlsx.utils.json_to_sheet(items);
  xlsx.utils.book_append_sheet(newWorkbook, sheet, category.slice(0, 31)); // Excel ma limit 31 znaków na nazwę arkusza
});

// Zapisz nowy plik
xlsx.writeFile(newWorkbook, outputFile);
console.log(`Zapisano jako: ${outputFile}`);
