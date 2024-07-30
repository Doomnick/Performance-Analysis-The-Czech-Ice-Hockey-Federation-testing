This script in R language processes data from the Monark Wingate device, MetaLyzer 3B R3 spiroergometry device, and anthropometric measurements. It summarizes the results, compares them with previous measurements and reference values, generates plots for each measurement, and creates a final radar chart depicting the athlete's performance profile. The script then exports a PDF file for each athlete trough R Markdown plugin and produces an XLSX file with the results for the entire tested group. It was developed to streamline and clarify the results of mass testing for youth and adults within the Czech Ice Hockey Federation.

Wingate_Spiro_radarchart.r is the main script. If you only want to generate results from spiroergometry, use Spiroergometry_reports.r instead.

Folder with Scripts (Wingate_Spiro_radarchart.r and export.rmd) must contain the following subfolders (names without diacritics, as written):

antropometrie - contains .xls files with anthropometric data (either one bulk file or separate files for each subject). In case of jump evaluations, the file includes a column "SJ" with numeric values.

databaze - the resulting Excel database with all previous measurements.

reporty - final PDF reports generated after running the script.

vymazat - temporary data that should be deleted after the script finishes.

vysledky - contains Excel export from the current measurement.

wingate - .txt exports from Wattbike (current measurements). Comparison folders can be located anywhere (the script will prompt for the file path).

spiro - .xlsx exports from spirometry.

Folders REPORTY and VYMAZAT should be empty before running the script.

Folders ANTROPOMETRIE, WINGATE, and possibly SPIRO must contain files with matching names (IDs) otherwise, files will not pair, and the script will display a warning.

Procedure:

1. Open the file Wingate_Spiro_radarchart.
2. On the 3rd line of the script, check the path to the folder with the script in the correct format (e.g., "C:/Users//Documents/Wingate+Spiro").
3. Run the script using the button on the top right of the code window.
4. A prompt for the spirometry report will appear. If you select YES, files from the spiro folder will be included in the report.
5. The next prompt is for comparison Wingate files; you can add two comparisons.
6. The following prompt is for adding the oldest comparison to the graph. If you select NO, the graph will not include the curve of the oldest comparison.
7. The script will then check the compatibility of all files (naming and count). If there is an error, an error message and description will be displayed in the console.
8. If you choose to continue with the export despite the error, incorrect data will not be evaluated, or may be partially evaluated. If you choose NO, the script will stop. Correct the files and start again by pressing the button.
9. If IDs and counts are correct, the script will proceed with generating the report.
10. The script generates reports into the folder.
11. The script will prompt to save the file to the database. Type A or N in the console and press Enter.
12. If A - the database will upload the last saved database file and add the current measurement to it, then export a new file.
13. Delete everything in the vymazat folder.

contact: do.kolinger@gmai****.com



