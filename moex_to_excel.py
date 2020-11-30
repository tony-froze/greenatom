import os
import sys

import datetime
import smtplib
import string
from email.message import EmailMessage
from itertools import product

import pymorphy2
import requests
import win32com.client as win32
from lxml.html import fromstring

BASE_URL = 'https://www.moex.com/ru/derivatives/currency-rate.aspx'


def get_all_page(url, params):
    try:
        response = requests.get(url, params=params)
    except requests.exceptions.RequestException:
        print('An unexpected connection error has occurred')
        sys.exit()
    parsed_page = fromstring(response.text)
    return parsed_page


def parse_data(parsed_page):
    # Get date and exchange rate(main clearing value) for it.
    dates = map(lambda arg: arg.text_content(), parsed_page.xpath('//tr[@class]/td[position()=1]'))
    rates = map(lambda arg: arg.text_content(), parsed_page.xpath('//tr[@class]/td[position()=3]'))
    data = zip(dates, rates)
    data = [get_table_cells(raw_row) for raw_row in data if '-' not in raw_row]
    data = take_cells_for_current_month(data)
    get_change_per_day(data)
    return [item for item in data if len(item) == 3]  # Only rows for the current month are returned.


def get_table_cells(raw_row):
    try:
        date = datetime.datetime.strptime(raw_row[0], '%d.%m.%Y')
        currency_rate = float(raw_row[1].replace(',', '.'))
        return [date, currency_rate]  # A list is more convenient for further work then a tuple.
    except ValueError:
        print('Unsupported data format')
        return [datetime.datetime.strptime('01.11.2020', '%d.%m.%Y'), 1.00]


def take_cells_for_current_month(data):
    current_month = datetime.datetime.now().month
    data_for_current_month = []
    for item in data:
        data_for_current_month.append(item)
        # I've left one row from the previous month to calculate rate change per day properly.
        if item[0].month < current_month:
            break
    return data_for_current_month


def get_change_per_day(data):
    for index in range(len(data)):
        try:
            change = data[index][1] - data[index + 1][1]
            data[index].append(change)
        except IndexError:
            pass


def create_table(data_to_write, headers, app):
    workbook = app.Workbooks.Add()
    worksheet = workbook.Worksheets.Add()
    worksheet.Name = 'USD and EUR to RUB rate'
    worksheet.Range('A1:G1').Value = headers
    rows = 1

    for table_row in data_to_write:
        rows += 1
        for i in range(len(table_row)):
            cell = worksheet.Cells(rows, i + 1)
            cell.Value = table_row[i]
        worksheet.Cells(rows, 7).Value = f'=E{rows}/B{rows}'

    for cols in (f'B2:C{rows}', f'E2:F{rows}'):
        # Cells look exactly the same as with currency format, but Excel recognizes them as custom.
        worksheet.Range(cols).NumberFormat = '_-* # ##0,00 ₽_-;-* # ##0,00 ₽_-;_-* ""-""?? ₽_-;_-@_-'

    for col in (f'A2:A{rows}', f'D2:D{rows}'):
        # The only way I've found to set date format readable for my Excel was to use russian letters.
        worksheet.Range(col).NumberFormat = 'ДД.ММ.ГГГГ'
    worksheet.Columns("A:G").AutoFit()

    alert_message = check_cells(app, len(headers), rows)

    return workbook, alert_message


def check_cells(app, num_columns, num_rows, has_header=True):
    """Checks all data has numeric value.

    This function isn't necessary here, because all data I put to  the table is already converted to float. Also,
    the way this check was done is the one big crunch. But it is the only way I found to check how Excel itself
    sees data in cells. I tried to add data validation before actually writing information to cell, but is doesn't
    work. As a  result I got the table with incorrect data, but with preset validation rule.
    """
    message = 'Все данные числового типа.'
    app.Cells(1, num_columns+1).Select()
    selected = app.ActiveCell
    for column, row in product(string.ascii_uppercase[num_columns], range(1 + has_header, num_rows + 1 + has_header)):
        selected.Formula = f'=TYPE({column}{row})'
        if selected.Value != 1:
            message = 'Таблица может содержать некорректные данные.'
            break
    selected.Delete()
    return message


def create_description_msg(num_of_rows, message):
    morph = pymorphy2.MorphAnalyzer()
    right_declension = morph.parse('строка')[0].make_agree_with_number(num_of_rows).word
    return f'В отчете {num_of_rows} {right_declension}.{message}'


def send_email(file_name, mgs_text):
    msg = EmailMessage()

    from_address = input("Type sender's address:")
    email_password = input("Type password:")
    to_address = input("Type receiver's address:")

    msg['Subject'] = f'Отчет за {datetime.datetime.now().date()}'
    msg['From'] = from_address
    msg['To'] = to_address
    msg.set_content(mgs_text)

    with open(file_name, 'rb') as file:
        file_data = file.read()
        msg.add_attachment(file_data, maintype='application',
                           subtype='octet-stream', filename=file_name)

    with smtplib.SMTP_SSL('smtp.mail.ru', 465) as smtp:
        smtp.login(from_address, email_password)
        smtp.send_message(msg)


def main():
    usd_page = get_all_page(BASE_URL, params={'currency': 'USD_RUB'})
    rub_to_usd = parse_data(usd_page)
    eur_page = get_all_page(BASE_URL, params={'currency': 'EUR_RUB'})
    rub_to_eur = parse_data(eur_page)
    combined_data = [usd_data + eur_data for usd_data, eur_data in zip(rub_to_usd, rub_to_eur)]
    headers = ['USD_date', 'USD_rate', 'USD_change', 'EUR_date', 'EUR_rate', 'EUR_change', 'EUR_to_USD']
    excel_app = win32.gencache.EnsureDispatch('Excel.Application')
    wb, data_type_valid_msg = create_table(combined_data, headers, excel_app)
    path = os.getcwd()
    filename = f'report_{datetime.datetime.now().strftime("%Y-%m-%d_%H.%M")}.xlsx'
    wb.SaveAs(f'{path}\\{filename}')
    wb.Close()
    excel_app.Quit()

    description_msg = create_description_msg(len(combined_data), data_type_valid_msg)
    print(description_msg)
    send_email(filename, description_msg)


if __name__ == '__main__':
    main()
