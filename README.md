# T-Shirts Organizer

This project is a tool to organize t-shirt orders for a sports event. It reads data from a Google Sheets spreadsheet, processes it to generate CSV files, and updates the spreadsheet with the processed information.

## Dependencies

To run this project, you'll need the following dependencies:

- `pandas`: A powerful data manipulation library.
- `gspread`: A Python API for Google Sheets.
- `gspread_dataframe`: An extension to gspread for working with Pandas DataFrames in Google Sheets.

You can install these dependencies using pip:

```bash
pip install pandas gspread gspread_dataframe
```

## Usage

1. Ensure you have a Google Sheets spreadsheet with the necessary columns: "Carimbo de data/hora", "Endereço de e-mail", "Nome completo", and columns for different t-shirt models and customizations.
2. Obtain the JSON file for the Google Service Account and provide its path in the code.
3. Update the URL of the Google Sheets spreadsheet in the code.
4. Run the script.
5. The script will update the spreadsheet and generate CSV files with the processed data.

## Important Links

- [Pandas Documentation](https://pandas.pydata.org/docs/)
- [gspread Documentation](https://gspread.readthedocs.io/en/latest/)
- [gspread_dataframe Documentation](https://github.com/robin900/gspread-dataframe)

