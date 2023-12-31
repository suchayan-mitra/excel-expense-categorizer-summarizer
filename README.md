# Excel Expense Summary Project
## Author: Suchayan Mitra

## Description
This project contains VBA scripts used to categorize transactions and generate a summary of expenses within Excel. It creates Pivot Tables and charts to visualize the data.

## Installation
To install, download the `.bas` files from this repository and import them into your Excel workbook:

1. Open Excel and press `Alt + F11` to open the VBA editor.
2. Right-click on any existing VBA project where you want to add the code.
3. Choose `Import File...` and select the `.bas` file.

## Usage
To use the scripts, (1) import the bank or credit card transactions in the simple format in the sheet "INPUT" as Date, Description, Debit, Credit, and Category. Leave the Category blank as it can be automatically categorized based on the mapping done in the code. Note: Update the mapping (vendor name pattern) in the code to categorize programmatically. (2) Run the `CategorizeTransactions` subroutine to categorize your transactions and then `SummarizeAndChartExpenses` to generate the summary and charts.

## Contributing
Contributions to this project are welcome. You can fork the project, make changes, and create a pull request to the main branch.

## License
This project is released under the MIT License. See the `LICENSE` file for details.

## Contact
Feel free to contact me for any questions or feedback.

