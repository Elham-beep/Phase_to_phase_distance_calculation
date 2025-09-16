#this script was used for debugging the calculation script, you can use it for counting sheets in any workbooks
#just provide the path below and it prints in the terminal
import pandas as pd

xl = pd.ExcelFile(r"path to the xlsx workbook you want to counts its sheets")
sheet_count = len(xl.sheet_names)

print(f"Number of sheets: {sheet_count}")
