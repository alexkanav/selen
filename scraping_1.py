import httplib2
from oauth2client.service_account import ServiceAccountCredentials
import googleapiclient.discovery
from selenium import webdriver
from bs4 import BeautifulSoup
import logging
from config import *



logging.basicConfig(level=logging.INFO, filename="main.log",filemode="a",
                    format="%(asctime)s %(levelname)s %(message)s")

CREDENTIALS_FILE = 'trade_key.json'
spredsheets_id = google_sheet_id
credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets',
                                                                                  'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = googleapiclient.discovery.build('sheets', 'v4', http=httpAuth)


def get_stock_name_from_google():
    values = service.spreadsheets().values().get(
        spreadsheetId=spredsheets_id,
        range=google_sheet_read,
        majorDimension='ROWS'
    ).execute()
    return values.get('values')


def write_to_google_sheet(stocks_ch):
    values = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spredsheets_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": start_cell_for_write,
                 "majorDimension": 'ROWS',
                 "values": stocks_ch}
            ]
        }
    ).execute()


def scraping(stock_name, driver):
    stock_ch = []
    try:
        driver.get(f'https://finance.yahoo.com/quote/{stock_name}/analysis/')
        driver.implicitly_wait(20)
        soup = BeautifulSoup(driver.page_source, "html.parser")

        earnings_estimate = soup.find('h3', string='Earnings Estimate')
        source1 = earnings_estimate.parent.find_next_sibling()
        earnings_avg_estimate = source1.find('td', string='Avg. Estimate')
        str1 = get_data(earnings_avg_estimate)
        year_ago_eps = source1.find('td', string='Year Ago EPS')
        str2 = get_data(year_ago_eps)

        revenue_estimate = soup.find('h3', string='Revenue Estimate')
        source2 = revenue_estimate.parent.find_next_sibling()
        revenue_avg_estimate = source2.find('td', string='Avg. Estimate')
        str3 = get_data(revenue_avg_estimate)
        sales_growth = source2.find('td', string='Sales Growth (year/est)')
        str4 = get_data(sales_growth)

        earnings_history = soup.find('h3', string='Earnings History')
        source3 = earnings_history.parent.find_next_sibling()
        eps_est = source3.find('td', string='EPS Est.')
        str5 = get_data(eps_est)
        eps_actual = source3.find('td', string='EPS Actual')
        str6 = get_data(eps_actual)
        difference = source3.find('td', string='Difference')
        str7 = get_data(difference)
        surprise = source3.find('td', string='Surprise %')
        str8 = get_data(surprise)

        growth_estimates = soup.find('h3', string='Growth Estimates')
        source4 = growth_estimates.parent.find_next_sibling()
        current_qtr = source4.find('td', string='Current Qtr.')
        str9 = get_data(current_qtr)
        next_qtr = source4.find('td', string='Next Qtr.')
        str10 = get_data(next_qtr)
        current_year = source4.find('td', string='Current Year')
        str11 = get_data(current_year)
        next_year = source4.find('td', string='Next Year')
        str12 = get_data(next_year)
        next_5_years = source4.find('td', string='Next 5 Years (per annum)')
        str13 = get_data(next_5_years)
        past_5_years = source4.find('td', string='Past 5 Years (per annum)')
        str14 = get_data(past_5_years)

        stock_ch = [
            str2[0], str1[0], str2[1], str1[1], str2[2], str1[2], str2[3], str1[3],
            str4[0], str3[0], str4[1], str3[1], str4[2], str3[2], str4[3], str3[3],
            str5[0], str6[0], str7[0], str8[0], str5[1], str6[1], str7[1], str8[1],
            str5[2], str6[2], str7[2], str8[2], str5[3], str6[3], str7[3], str8[3],
            str9[0], str10[0], str11[0], str12[0], str13[0], str14[0], str9[1], str10[1], str11[1], str12[1], str13[1], str14[1],
            str9[2], str10[2], str11[2], str12[2], str13[2], str14[2], str9[3], str10[3], str11[3], str12[3], str13[3], str14[3]
        ]
    except Exception as err:
        print('Error: ', err)
        stock_ch = [
            '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-',
            '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-',
            '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-', '-',
            '-', '-', '-', '-', '-', '-', '-', '-'
        ]

    return stock_ch


def get_data(item):
    vals = []
    for i in range(4):
        item = item.find_next_sibling()
        vals.append(item.text)

    return vals


def main():
    stock_names = get_stock_name_from_google()
    stocks_ch = []
    if browser == 'Chrome':
        options = webdriver.ChromeOptions()
        options.add_argument('--headless=new')  # work in background mode
        options.add_argument("start-maximized")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        options.add_argument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.103 Safari/537.36")
        options.add_argument("--disable-blink-features=AutomationControlled")
        driver = webdriver.Chrome(options=options)
    elif browser == 'FireFox':
        driver = webdriver.Firefox()
        driver.maximize_window()
    else:
        driver = None

    for stock_name in stock_names:
        stocks_ch.append(scraping(stock_name[0], driver))
    driver.close()
    write_to_google_sheet(stocks_ch)


if __name__ == "__main__":
    main()
