import xlwings as xw
from bs4 import BeautifulSoup
import urllib3
import urllib.parse as urlparse
import re


class Page:
    pass


wb = xw.Book.caller()
#wb = xw.books.open('keyword_links.xlsm')
sheetname = 'Sheet1'
ws = wb.sheets(sheetname)
checked_urls = []
pages = []
new_book = xw.Book()
wb.activate()


def get_page(domain, url, keywords, d_num):
    if url in checked_urls:
        return
    checked_urls.append(url)
    if len(checked_urls) > 500:
        return
    print(checked_urls)

    http = urllib3.PoolManager()
    html = http.request('GET', url)
    soup = BeautifulSoup(html.data, 'html.parser')
    links = filter(lambda link: link.find('tel:') != 0, [
                   a['href'] for a in soup.find_all('a', href=True)])
    links2 = []
    for link in links:
        if link.find('http') != 0:
            links2.append(urlparse.urljoin(domain, link))
        else:
            links2.append(link)
    links2 = filter(lambda link: link.find(domain) == 0, links2)
    links2 = [link.split('?')[0] for link in links2]
    links2 = [link.split('#')[0] for link in links2]
    links2 = [link[:-1] if link[-1] == '/' else link for link in links2]
    links2 = filter(lambda link: '.' not in link or '.html' in link, links2)
    title = soup.find('title').text if soup.find('title') is not None else ''
    description_tag = soup.find('meta', attrs={'name': 'description'})
    description = description_tag['content'] if description_tag is not None else ''

    for tag in soup.find_all('a'):
        tag.decompose()
    hit_words = []
    for keyword in keywords:
        if soup.find(text=re.compile(keyword)):
            hit_words.append(keyword)

    if len(hit_words) > 0:
        page = Page()
        page.url = url
        page.hit_words = ','.join(hit_words)
        page.title = title
        page.description = description
        pages.append(page)
        ws.range(2+d_num, 4).value = len(pages)

    for link in links2:
        get_page(domain, link, keywords, d_num)


def main():
    keywords = []
    for row in range(2, 100):
        keyword = ws.range(row, 3).value
        if keyword != None:
            keywords.append(keyword)
        else:
            break

    domains = []
    for row in range(2, 100):
        domain = ws.range(row, 2).value
        if domain != None:
            domains.append(domain)
        else:
            break

    for d_num, domain in enumerate(domains):
        print(domain)
        global checked_urls
        global pages
        checked_urls = []
        pages = []
        get_page(domain, domain, keywords, d_num)

        if d_num == 0:
            sheet = new_book.sheets('Sheet1')
        else:
            sheet = new_book.sheets.add()
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


# if __name__ == '__main__':
#    main()
