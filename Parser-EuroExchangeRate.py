import requests
from bs4 import BeautifulSoup
import pandas
import datetime

urlBase = 'https://index.minfin.com.ua/exchange/archive/nbu/curr/'
dictOfRates = {}


# This function searches the pages for the value of the EURO exchange rate
def finedEURO(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, 'lxml')
    quotes = soup.find_all('td', class_='r1')

    for i in range(len(quotes)):
        if 'Евро' in quotes[i]:
            return quotes[i+1].text


# Set range of months to process
year = int(input("Enter the year in YYYY format: "))
startMonth = int(input("Enter the number of the first month: "))
lastMonth = int(input("Enter the number of the last month: ")) + 1

# The full link to the page containing the exchange rate looks like this:
# https://index.minfin.com.ua/exchange/archive/nbu/curr/2023-02-01/
# Here dates are added to the source url
for month in range(startMonth, lastMonth):
    monthForOutput = datetime.datetime(2018, month, 1).strftime("%B")  # Custom date template to get the month
    print('Working...',  monthForOutput)
    for day in range(1, 32):
        try:
            date = datetime.date(year, month, day)
            urlFull = urlBase + str(date)
            dictOfRates[str(date)] = [finedEURO(urlFull)]
        except ValueError:
            break  # Way to handle months with less than 31 days

# Saving the result to a file
df = pandas.DataFrame(list(dictOfRates.values()),
                      index=list(dictOfRates.keys()),
                      columns=['EURO'])

resultIsSaved = False

while not resultIsSaved:
    try:
        df.to_excel('result.xlsx')
        resultIsSaved = True
    except PermissionError:
        input('Please close result.xlsx file and press ENTER')

print('DONE. Please see the result in the file result.xlsx')
