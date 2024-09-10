# Excel VBA Project for Data Management and Automation
Overview

This project features an Excel workbook containing uncleaned data from Nester's Market periodic Gross Profit reports, enhanced with VBA macros designed to automate data management, cleaning, and processing tasks. The key functions include cleaning headers, unmerging and splitting cells, managing inventory data, and consolidating information across multiple worksheets.

## Features
- CleanHeadersAndTitles: Automatically identifies and removes empty rows, redundant headers, and titles, ensuring a clean dataset for further analysis.
- InsertColumnsAndHeaders: Adds additional columns, such as 'Store Name' and 'Period,' to improve data organization.
- UnmergeAndSplitCells: Unmerges cells and splits text based on specific criteria, improving data readability.
- MoveTotalToAdjacentCell: Moves text from "Total" cells into adjacent cells for cleaner report formatting.
- ProcessAndEnsureAdjacentCellIsBlank: Populates store names upwards based on "Total" rows and ensures adjacent cells are left blank when necessary.
- UpdatePeriod: Automatically extracts the period from worksheet names and inserts it into the appropriate column.
- ClearFormatAndAdjustAllSheets: Removes all formatting and auto-adjusts column and row sizes for better visibility.
- InsertOrReplaceClearTableInAllSheets: Inserts or replaces tables across all sheets to ensure consistent formatting.
- ConsolidateDataAndRename: Merges data from multiple sheets into a "Master" sheet and renames it to "Consolidated Data.
