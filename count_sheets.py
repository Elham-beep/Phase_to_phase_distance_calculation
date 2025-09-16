import pandas as pd

xl = pd.ExcelFile(r"C:\Users\bimax\DC\ACCDocs\Axpo Grid AG\DEMO_AXPO_Leitungen\Project Files\Grid 4.0 - PLS Distances Development\PLS_CADD_tests_Elham\Cond_10C_TR1730a002_003_processed.xlsx")
sheet_count = len(xl.sheet_names)

print(f"Number of sheets: {sheet_count}")