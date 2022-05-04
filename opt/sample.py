import math
import sys

import openpyxl


def main():
  # val = float(sys.argv[1])
  wb = openpyxl.load_workbook('/root/file/Export Trade History-2021-12-04 15_10_17.xlsx')
  sheet = wb['sheet1']
  value = sheet.cell(row=1, column=2).value
  wb.close()
  print(value)

if __name__ == "__main__":
  main()
