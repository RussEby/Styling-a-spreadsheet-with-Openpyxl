# import the needed items
from openpyxl import load_workbook
from openpyxl.packaging.core import DocumentProperties


def main(old_file, new_file):
    # load the plain file
    wb = load_workbook(filename=old_file)

    # Set up some Excel Properties
    wb.properties = DocumentProperties(
        creator="Russ", title="Country List", lastModifiedBy="OpenPyXL"
    )

    # select the active page
    sheet = wb.active

    # Provide a good name for the sheet
    sheet.title = "Country Data"

    #
    # This cleaning assumes all cells are strings
    #
    # iterate over every column
    for column_cells in sheet.columns:
        # iterate over evey cell in the column
        for cell in column_cells:
            # if the cell has data
            if cell.value:
                # strip spaces at the start and end
                cell.value = cell.value.strip()

    # save the OpenPyXL Workbook
    wb.save(new_file)


if __name__ == "__main__":
    main(old_file="data/CSV_load.xlsx", new_file="data/cleaned.xlsx")
