import openpyxl
from nbkits.xlstyler import xlstyle
from pathlib import Path


def main():
    wb = openpyxl.Workbook()
    (
        xlstyle(wb.active)
        .column_width(3, col=range(2, 30))
        .row_height(10, row=range(2, 30))
        #
        .border(col=range(2, 7), row=range(2, 6), sides="outside", ls="thick")
        #
        .border(
            col=range(2, 7),
            row=range(8, 12),
            sides="inside",
            ls="medium",
            c="r",
        )
        #
        .border(col=range(2, 7), row=range(14, 18), sides="outside", ls="thick")
        .border(
            col=range(2, 7),
            row=range(14, 18),
            sides="inside",
            ls="medium",
            c="r",
        )
        #
        .border(
            col=range(2, 7),
            row=range(20, 24),
            sides="inside",
            ls="medium",
            c="b",
        )
        .border(col=range(2, 7), row=range(20, 24), sides="outside", ls="thick")
        .border(
            col=range(9, 13),
            row=range(2, 6),
            t="thick",
            b="double",
            ls="thin",
            c="xkcd:red",
        )
    )

    test_dir = Path(__file__).parent
    wb.save(test_dir / "test_borders.xlsx")


if __name__ == "__main__":
    main()
