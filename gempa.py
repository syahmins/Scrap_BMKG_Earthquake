# This script will take data from BMKG using their official API.
# This is to scrape 15 earthquake with magnitude 5.0+ in Indonesia and save it in Excel.
# BMKG website: bmkg.go.id

import requests
import json
import xlsxwriter

# JSON location
url = 'https://data.bmkg.go.id/DataMKG/TEWS/gempaterkini.json'
results = requests.get(url).json()


def buat_excel():
    # Make a variable for column title
    judul_kolom = ['No.', 'Jam', 'Koordinat', 'Magnitude', 'Kedalaman', 'Wilayah', 'Potensi Tsunami']

    # Create an Excel file and a workbook
    workbook = xlsxwriter.Workbook('latest_earthquake_idn.xlsx')
    worksheet = workbook.add_worksheet()
    row = 3

    # write the column title
    for col, title in enumerate(judul_kolom):
        worksheet.write(row, col, title)

    # earthquake date
    worksheet.write(0, 0, f"{results['Infogempa']['gempa'][0]['Tanggal']}")

    # reset numbers and rows
    nomor = 0
    row = 4

    for i in range(15):
        worksheet.write(row, 0, nomor)  # first column (column A3)
        worksheet.write(row, 1, f"{results['Infogempa']['gempa'][nomor]['Jam']}")  # 2nd column (column A4)
        worksheet.write(row, 2, f"{results['Infogempa']['gempa'][nomor]['Coordinates']}")  # 3rd column (column A5)
        worksheet.write(row, 3, f"{results['Infogempa']['gempa'][nomor]['Magnitude']}")  # 4th column (column A6)
        worksheet.write(row, 4, f"{results['Infogempa']['gempa'][nomor]['Kedalaman']}")  # 5th column (column A7)
        worksheet.write(row, 5, f"{results['Infogempa']['gempa'][nomor]['Wilayah']}")  # 6th column (column A8)
        worksheet.write(row, 6, f"{results['Infogempa']['gempa'][nomor]['Potensi']}")  # 7th column (column A9)

        nomor += 1
        row += 1

    # done
    workbook.close()
    print('Data sudah selesai diekspor ke Excel')


if __name__ == '__main__':
    buat_excel()
