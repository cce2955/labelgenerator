import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter

def get_user_input():
    data = []
    for row in range(22):  # Adjust for the new number of rows
        row_data = []
        for col in range(8):
            user_input = input(f"Enter input for cell ({row+1}, {col+1}) (type 'exit' to stop): ")
            if user_input.lower() == 'exit':
                while len(row_data) < 8:  # Fill the remaining cells in the row with empty strings
                    row_data.append('')
                data.append(row_data)
                while len(data) < 22:  # Fill the remaining rows with empty strings
                    data.append([''] * 8)
                return data
            row_data.append(user_input)
        data.append(row_data)
    return data

def create_excel_sheet(data):
    wb = openpyxl.Workbook()
    ws = wb.active

    # Set the column width and row height to approximate 1 inch wide and 0.5 inch high cells
    one_inch_points = 72  # 1 inch = 72 points
    col_width = 10  # Width in Excel units, adjust if necessary
    row_height = one_inch_points / 2  # Half inch in points

    for col in range(1, 9):  # Excel columns start from 1
        ws.column_dimensions[get_column_letter(col)].width = col_width

    for row in range(1, 23):  # Adjust for the new number of rows
        ws.row_dimensions[row].height = row_height

    # Define a border style for the cutting guide
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    max_font_size = 14
    min_font_size = 6

    # Fill the cells with data and apply borders and dynamic font size
    for r_idx, row_data in enumerate(data, start=1):
        for c_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.border = thin_border
            
            # Calculate appropriate font size based on length of the input
            if len(value) <= 10:
                font_size = max_font_size
            else:
                font_size = max(max_font_size - len(value) // 2, min_font_size)
            
            if font_size < min_font_size:
                cell.value = value[:int((max_font_size - min_font_size) * 2)]  # Truncate to fit within the limits
                font_size = min_font_size

            cell.font = Font(size=font_size)
            cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

    # Set print options
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.print_area = 'A1:H22'  # Ensure the print area covers the new grid size
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    # Save the workbook
    wb.save("user_input_grid.xlsx")
    print("Excel sheet created as 'user_input_grid.xlsx'.")

def main():
    print("You will be prompted to enter inputs for an 8x22 grid (1 inch wide, 0.5 inch high cells).")
    data = get_user_input()
    create_excel_sheet(data)

if __name__ == "__main__":
    main()
