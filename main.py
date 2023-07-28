""" Pyxelize"""

import os
import sys
import json
import requests
import xlwt

from dotenv import load_dotenv

load_dotenv()

RAPIDAPI_KEY = os.getenv('RAPIDAPI_KEY')
RAPIDAPI_HOST = os.getenv('RAPIDAPI_HOST')

endpoint = "genres"
url = f"https://{RAPIDAPI_HOST}/v2/{endpoint}"

headers = {
	"X-RapidAPI-Key": RAPIDAPI_KEY,
	"X-RapidAPI-Host": RAPIDAPI_HOST
}


def main(args=sys.argv[:1]):
    """ Main program """

    response = requests.get(url, headers=headers)
    json_result = response.json()["result"]

    wb = xlwt.Workbook()
    sheet = wb.add_sheet("Genres")

    sheet.write(0, 0, "Genre")
    sheet.write(0, 1, "Nombre")

    nb_of_movies_fmt = xlwt.easyxf(
        "pattern: pattern solid, fore_color blue;" \
        "font: color white;" \
        "align: vert centre, horiz centre"
    )

    index = 1
    for nb_of_movies, gender_name in dict(json_result).items():
        sheet.write(index, 0, gender_name)
        sheet.write(index, 1, nb_of_movies, nb_of_movies_fmt)
        index += 1

    sheet.col(0).width = 30 * 256
    sheet.col(1).width = 10 * 256

    wb.save("data/movies.xls")
    print("Sauvegardé movies.xls")

    return 0


if __name__ == "__main__":
    sys.exit(main())
