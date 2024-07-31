# from openpyxl import load_workbook
import openpyxl
from nbkits.xlstyler import xlstyler


def main():
    # 创建一个新的工作簿
    wb = openpyxl.Workbook()
    ws = wb.active

    (
        xlstyler(ws)
        .column_width(5, cols=range(2, 10))
        .row_height(10, rows=range(2, 20))
        #
        .border(cols=range(2, 7), rows=range(2, 6), sides="outside", ls="thick")
        #
        .border(
            cols=range(2, 7),
            rows=range(8, 12),
            sides="inside",
            ls="medium",
            c="ff0000",
        )
        #
        .border(cols=range(2, 7), rows=range(14, 18), sides="outside", ls="thick")
        .border(
            cols=range(2, 7),
            rows=range(14, 18),
            sides="inside",
            ls="medium",
            c="ff0000",
        )
        #
        .border(
            cols=range(2, 7),
            rows=range(20, 24),
            sides="inside",
            ls="medium",
            c="ff0000",
        )
        .border(cols=range(2, 7), rows=range(20, 24), sides="outside", ls="thick")
    )

    wb.save("example.xlsx")


if __name__ == "__main__":
    main()
