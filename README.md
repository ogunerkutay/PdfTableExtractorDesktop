# PDF Table Extractor Desktop

This repository contains a C# port of the original PDFTableExtractor with the developer's permission. The application extracts tables from PDF files and converts them into Excel (XLSX) format.

## Downloading & Installing

1. Go to the **Releases** section and download the latest installer file.
2. Run the installer (avoid installing to `ProgramFiles`).
3. A **PDFTableExtractor** shortcut will appear on the Desktop.

## Usage with Example

1. Drag & drop one or more PDFs onto the Desktop shortcut.
2. Alternatively, right-click on the PDF and select the extract option (must be enabled in settings).
3. A command prompt will appear, printing information about the processing.
4. XLSX files will be created in the same directory where the PDFs were located.

For customizing output, check out the Settings wiki page.

## Example Input & Output

![Image](https://github.com/user-attachments/assets/454de69c-dee0-4dee-8cf1-f67821d813a4)

## Settings

To bring up the settings menu, start the desktop icon normally.

- **Keep pages with rows/columns**: Skip exporting all pages/sheets that don't meet the criteria.
- **Skip empty rows/columns**: Different options for choosing row/column skipping methods.
- **Page naming strategy**: How to name pages/sheets in the Excel file.
- **Autosize columns**: Resizes created columns before saving.
- **Parallel file processing**: Enables processing multiple PDFs at the same time.
- **Context menu**: When turned on, an extraction option appears in the right-click menu of PDF files.

![Image](https://github.com/user-attachments/assets/96a73a30-7188-4f45-8b08-be4006b9c91c)

## Updating

1. When a new version is available, a message will appear in the console saying that the local version is out of date.
2. Go to the **Releases** section.
3. Download the new installer.
4. Uninstall the old version.
5. Install the new version.

## Reporting Bugs

1. Create a new issue with a descriptive title.
2. Try to include more information, e.g., the PDF you tried to extract (if you're allowed to), your settings, `error.txt`.
3. If the expected output is wrong, demonstrate what the expected output would be and what the output of the app was.
4. When a program error occurs, a file named `error.txt` gets created in the directory of the application.

## Requesting New Features

1. Create an issue describing the feature/filter you need, giving it a descriptive name.
2. Write a short description of what the feature/filter would do.
3. Post screenshots of the input and the expected output.

---

This C# port of the original PDFTableExtractor was created with the developer's permission.
