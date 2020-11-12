import excel

excel.gen_csv("Test.csv")
excel.print_csv("Test.csv")

excel.csv_to_xlsx("Test.csv", "Test.xlsx")
excel.read_xlsx("Test.xlsx")

excel.csv_to_xlsx_table("Test.csv", "Test2.xlsx")
excel.read_xlsx("Test2.xlsx")
