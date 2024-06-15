# User Input Excel Sheet Generator

This project is a Python script that creates an Excel sheet based on user input. The Excel sheet is formatted with specific dimensions and includes dynamic font sizing for cell content.

## Features

- Prompts the user to enter inputs for a grid of 8 columns by 22 rows.
- Dynamically adjusts font size based on the length of the input.
- Formats cells with borders and sets specific column widths and row heights.
- Saves the Excel sheet as `user_input_grid.xlsx`.

## Requirements

- Python 3.x
- `openpyxl` library

## Installation

1. Clone the repository or download the script file.
2. Navigate to the directory containing the script.
3. Install the required library using the following command:
   '''sh
   pip install -r requirements.txt
   '''

## Usage

1. Run the script using Python:
   '''sh
   python main.py
   '''
2. You will be prompted to enter inputs for each cell in an 8x22 grid. Type your input and press Enter to move to the next cell.
3. To stop entering inputs before the grid is completely filled, type `exit`.
4. The script will generate an Excel file named `user_input_grid.xlsx` in the same directory.

## Example

'''sh
$ python main.py
You will be prompted to enter inputs for an 8x22 grid (1 inch wide, 0.5 inch high cells).
Enter input for cell (1, 1) (type 'exit' to stop): Hello
Enter input for cell (1, 2) (type 'exit' to stop): World
...
Enter input for cell (2, 1) (type 'exit' to stop): exit
Excel sheet created as 'user_input_grid.xlsx'.
'''

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/) for providing the tools to handle Excel files in Python.
