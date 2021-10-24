import xlwings as xw
from bs4 import BeautifulSoup
import urllib3


def get_links(domain, url):
    http = urllib3.PoolManager()
    html = http.request('GET', url)
    soup = BeautifulSoup(html.data, 'html.parser')
    tags = soup.find_all('a')
    links = []
    for tag in tags:
        link = tag.get('href')
        print(link)
        if link in domain:
            links.append(link)
    return links


def main():

    wb = xw.Book.caller()
    sheetname = 'Sheet1'
    ws = wb.sheets(sheetname)

    #domain = ws.range(2, 2).value
    domain = 'http://52.196.27.1/genki-hp/'
    print(domain)
    links = get_links(domain, domain)
    for i, link in enumerate(links):
        #ws.range(i+2, 3).value = link
        print(link)

    #http = urllib3.PoolManager()
    #html = http.request('GET', url)
    #soup = BeautifulSoup(html.data, 'html.parser')
    #title = soup.title.string
    #ws.range(2, 3).value = title
