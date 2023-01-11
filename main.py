from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font

data = {
    "Student 1": {
        "math": 75,
        "science": 78,
        "history": 90
    },
    "Student 2": {
        "math": 79,
        "science": 71,
        "history": 72
    },
    "Student 3": {
        "math": 89,
        "science": 87,
        "history": 72
    }
}

wb = Workbook()
ws = wb.active
ws.title = "Grades"

headings = ["Name"] + list(data["Student 1"].keys())
ws.append(headings)

for student in data:
    grades = list(data[student].values())
    ws.append([student] + grades)

for col in range(2, len(data["Student 1"]) + 2):
    char = get_column_letter(col)
    ws[char + "7"] = f"=SUM({char + '2'}:{char + '6'})/{len(data)}"

for col in range(1, len(data["Student 1"]) + 2):
    ws[get_column_letter(col) + "1"].font = Font(bold=True, color="0099CCFF")

wb.save("NewGrades.xlsx")