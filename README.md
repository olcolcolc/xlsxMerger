# XLSX Sort, Merge & Export Helper

This project started as a simple Node.js script to work with `.xlsx` files — sorting, filtering, and generating new sheets. It helps automate working with Excel data.

I developed this project to assist my boyfriend with organizing data for his PhD. For example, sorting sources by categories, merging data from different files, and exporting the results into new `.xlsx` files.

## Features

1. **Sorting**: Sorts data alphabetically based on a specific column.
2. **Merging**: Merges data from multiple Excel files, matching entries from different sources.
3. **Exporting**: Saves the results into a new `.xlsx` file after sorting or merging.

## How It Works

### Sort
1. The script reads an Excel file.
2. It sorts the data based on a chosen field (e.g., categories).
3. It generates a new file with the sorted data.

### Merge
1. It reads two or more Excel files.
2. It matches entries between files based on a common field (e.g., titles).
3. It generates a new file with the merged data.

## Requirements

- Node.js
- `xlsx` package

## Installation

1. Install dependencies:
   ```bash
   npm install xlsx
2. Place your Excel files in the project folder.

3. Run the script:
    ```bash
   node sort.js      # For sorting
    node merge.js     # For merging

After execution, you'll receive new .xlsx files containing the sorted or merged data.

### Plans for Future Development
- Adding more filtering options based on different criteria.
- Grouping data by categories.
- Creating a simple UI for executing actions without coding.
