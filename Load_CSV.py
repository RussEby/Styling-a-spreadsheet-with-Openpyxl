# import the needed items
import csv
import openpyxl


def main(csv_file, xslx_file):
    # Create a new OpenPyXL Workbook
    wb = openpyxl.Workbook()

    # Select the active sheet
    ws = wb.active

    # open the csv file
    with open(csv_file) as f:
        # create the reader object of the data in the CSV
        reader = csv.reader(f, delimiter=",")

        # load the data, row by row
        for row in reader:
            ws.append(row)

    # save the OpenPyXL Workbook
    wb.save(xslx_file)


if __name__ == "__main__":
    main("data/countries of the world.csv", "data/CSV_load.xlsx")
