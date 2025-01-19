const xlsx = require("xlsx");
const path = require("path");

// Ścieżki do plików
const file1Path = path.resolve("design_for_disassembly.xlsx");
const file2Path = path.resolve("cite_score.xlsx");
const outputFilePath = path.resolve("merged_data.xlsx");

// Wczytaj oba pliki Excel
const file1 = xlsx.readFile(file1Path);
const file2 = xlsx.readFile(file2Path);

// Wczytaj dane z odpowiednich arkuszy
const sheet1 = xlsx.utils.sheet_to_json(file1.Sheets[file1.SheetNames[0]], { header: 1 }); // header: 1 traktuje wiersze jako tablice
const sheet2 = xlsx.utils.sheet_to_json(file2.Sheets[file2.SheetNames[0]]);

// Krok 1: Tworzenie obiektu titlesScoreObject
const secondRow = sheet1[1]; // Drugi wiersz zawiera nagłówki
const sourceIndex = secondRow.indexOf("source"); // Znalezienie indeksu kolumny "source"
console.log(sourceIndex)

if (sourceIndex === -1) {
    console.error('Kolumna "source" nie została znaleziona w pliku file1.');
    process.exit(1);
}

const titlesScoreObject = {};
sheet1.forEach((row) => { 
    const sourceTitle = row[sourceIndex];
    if (sourceTitle) {
        titlesScoreObject[sourceTitle] = ""; // Inicjalizuj z pustą wartością
    }
});

// Sprawdzenie liczby kluczy
const titlesCount = Object.keys(titlesScoreObject).length;
console.log(`Znaleziono ${titlesCount} kluczy`)

// Wyświetlenie kluczy w konsoli
console.log("TitlesScoreObject keys:", Object.keys(titlesScoreObject));

// Krok 2: Tworzenie obiektu sourceMap
const sourceMap = {};
sheet2.forEach((row) => {
    const sourceTitle = row["Source title"]; // Tytuł z pliku file2
    const score = row["CiteScore"]; // Ocena z pliku file2
    if (sourceTitle) {
        sourceMap[sourceTitle] = score;
    }
});

// Krok 3: Przypisywanie wartości do titlesScoreObject
Object.keys(titlesScoreObject).forEach((title) => {
    if (sourceMap[title]) {
        titlesScoreObject[title] = sourceMap[title]; // Przypisz wartość, jeśli klucz istnieje w sourceMap
    } else {
        titlesScoreObject[title] = "Not found"; // Przypisz "Not found", jeśli klucz nie istnieje
    }
});

// Wyświetlenie wynikowego obiektu titlesScoreObject w konsoli
console.log("Updated TitlesScoreObject:", titlesScoreObject);

// Krok 4: Tworzenie pliku XLSX
const mergedData = Object.entries(titlesScoreObject).map(([title, score]) => ({
    "Source Title": title,
    "CiteScore": score,
}));

const newWorkbook = xlsx.utils.book_new();
const newSheet = xlsx.utils.json_to_sheet(mergedData);
xlsx.utils.book_append_sheet(newWorkbook, newSheet, "Merged Data");
xlsx.writeFile(newWorkbook, outputFilePath);

console.log(`Plik został zapisany jako: ${outputFilePath}`);
