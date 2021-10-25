import xlwings as xw
from bs4 import BeautifulSoup
import urllib3
import urllib.parse as urlparse

checked_urls = []
pages = []


class Page:
    pass


def get_page(domain, url, keywords):
    if url in checked_urls:
        return
    checked_urls.append(url)
    if len(checked_urls) > 500:
        return

    http = urllib3.PoolManager()
    html = http.request('GET', url)
    soup = BeautifulSoup(html.data)
    links = filter(lambda link: link.find('tel:') != 0, [
                   a['href'] for a in soup.find_all('a', href=True)])
    links2 = []
    for link in links:
        if link.find('http') != 0:
            links2.append(urlparse.urljoin(domain, link))
        else:
            links2.append(link)
    links2 = filter(lambda link: link.find(domain), links2)
    title = soup.find('title').text if soup.find('title') is not None else ''
    print(soup.find('meta', attrs={'name': 'description'}))
    description_tag = soup.find_all('meta', attrs={'name': 'description'})[0]
    description = description_tag['content'] if description_tag is not None else ''

    hit_words = []
    for keyword in keywords:
        if soup.find(text=keyword):
            hit_words.append(keyword)

    if len(hit_words) > 0:
        page = Page()
        page.url = url
        page.hit_words = ','.join(hit_words)
        page.title = title
        page.description = description
        pages.append(page)

    for link in links:
        #get_page(domain, link, keywords)
        pass


def main():
    #wb = xw.Book.caller()
    wb = xw.books.open('keyword_links.xlsm')
    sheetname = 'Sheet1'
    ws = wb.sheets(sheetname)

    keywords = []
    for row in range(1, 100):
        keyword = ws.range(row, 3).value
        if keyword != None:
            keywords.append(keyword)
        else:
            break
    domain = ws.range(2, 2).value
    get_page(domain, domain, keywords)

    sheet = wb.sheets.add()
    sheet.range(1, 1).value = 'ドメイン'
    sheet.range(1, 2).value = domain
    sheet.range(3, 1).value = 'url'
    sheet.range(3, 2).value = 'title'
    sheet.range(3, 3).value = 'hit keywords'
    sheet.range(3, 4).value = 'description'
    for i, page in enumerate(pages):
        sheet.range(4+i, 1).value = page.url
        sheet.range(4+i, 2).value = page.title
        sheet.range(4+i, 3).value = page.hit_words
        sheet.range(4+i, 4).value = page.description


if __name__ == '__main__':
    main()
