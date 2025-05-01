const xlsx = require("xlsx");
const path = require("path");


// Ścieżki do plików
const file1Path = path.resolve("circular_economy.xlsx");
const file2Path = path.resolve("cite_score.xlsx");
const outputFilePath = path.resolve("merged_data.xlsx");

// Wczytaj oba pliki Excel
const file1 = xlsx.readFile(file1Path);
const file2 = xlsx.readFile(file2Path);

// Wczytaj dane z odpowiednich arkuszy
const sheet1 = xlsx.utils.sheet_to_json(file1.Sheets[file1.SheetNames[0]], { header: 1 }); // header: 1 traktuje wiersze jako tablice
const sheet2 = xlsx.utils.sheet_to_json(file2.Sheets[file2.SheetNames[0]]);

// Krok 1: Tworzenie obiektu titlesScoreObject
const secondRow = sheet1[0]; // Pierwszy wiersz zawiera nagłówki
// console.log("nagłówki", sheet1[0])

const sourceIndex = secondRow.indexOf("source"); // Znalezienie indeksu kolumny "source"
// console.log(sourceIndex)

if (sourceIndex === -1) {
    console.error('Kolumna "source" nie została znaleziona w pliku file1.');
    process.exit(1);
}

const titlesScoreObject = {};
sheet1.forEach((row, index) => {
    const sourceTitle = row[sourceIndex];
    if (sourceTitle) {
        const sourceTitleToLowerCase = sourceTitle.toLowerCase() // sformatowanie do małej litery
        titlesScoreObject[`${index}`] = { title: sourceTitleToLowerCase }; // unikalny klucz, a wartością jest tytuł
    }
});

// Sprawdzenie liczby tytułów
const titlesCount = Object.keys(titlesScoreObject).length;
console.log(`Znaleziono ${titlesCount} tytułów`)

// Wyświetlenie kluczy w konsoli
// console.log("TitlesScoreObject keys:", Object.keys(titlesScoreObject));

// Krok 2: Tworzenie obiektu sourceMap z pliku xlsx "city_score"
const sourceMap = {};
sheet2.forEach((row) => {
    const sourceTitle = row["Source title"]; // Tytuł z pliku file2
    const score = row["CiteScore"]; // Ocena z pliku file2
    if (sourceTitle) {
        const sourceTitleToLowerCase = sourceTitle.toLowerCase() // sformatowanie do małej litery
        sourceMap[sourceTitleToLowerCase] = score;
    }
});

// Krok 3: Przypisywanie wartości do titlesScoreObject - obiektu, który zawiera tytuł z file1 i odpowiednią ocenę z file2
Object.keys(titlesScoreObject).forEach((key) => {
    const title = titlesScoreObject[key].title;
    titlesScoreObject[key].CiteScore = sourceMap[title] || "Not found";
});
// Przypisz wartość, jeśli klucz istnieje w sourceMap
// Przypisz "Not found", jeśli klucz nie istnieje


// Wyświetlenie wynikowego obiektu titlesScoreObject w konsoli
// console.log("Updated TitlesScoreObject:", titlesScoreObject);

// Krok 4: Tworzenie pliku XLSX
const mergedData = Object.values(titlesScoreObject).map(({ title, CiteScore }) => ({
    "Source Title": title,
    "CiteScore": CiteScore,
}));


const newWorkbook = xlsx.utils.book_new();
const newSheet = xlsx.utils.json_to_sheet(mergedData);
xlsx.utils.book_append_sheet(newWorkbook, newSheet, "Merged Data");
xlsx.writeFile(newWorkbook, outputFilePath);

console.log(`Plik został zapisany jako: ${outputFilePath}`);
