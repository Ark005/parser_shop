import xlwt
from bs4 import BeautifulSoup
import requests

url = 'https://www.atrium.su/specialty/stores/'

def make_request():
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "lxml")
    shop_card = soup.find_all('div', {'class': 'item_title'})
    parsed_shop_cards = []
    for card in shop_card:
        title = card.find('a', {'class': 'departmentCard__title-FuW3f'}).text
        category = card.find('div', {'class': 'departmentCard__categories-fCtho'}).text
        stage = card.find('div', {'class': 'shop_card__bottom'}).find('a', {'class': 'link-show-map'}).find('span').text
        stage = stage.replace('На схеме','').strip()
        parsed_card = {'title': title, 'category': category, 'stage': stage}
        parsed_shop_cards.append(parsed_card)
    return parsed_shop_cards


def save_result(results):
    # Initialize a workbook
    book = xlwt.Workbook()

    # Add a sheet to the workbook
    sheet1 = book.add_sheet("Sheet1")

    # Loop over the rows and columns and fill in the values

    for index, card in enumerate(results):
        row = sheet1.row(index)
        row.write(0, card['title'])
        row.write(1, card['category'])
        row.write(2, card['stage'])

    # Save the result
    book.save("atrium.xls")


result = make_request()
save_result(result)