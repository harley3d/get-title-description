# 1. упаковываем урлы в массив
# 2. Идем по списку массива
# 3. Открываем урл массива
# 4. Копируем title
# 5. Копируем description
# 6. Копируем keywords
# 7. Открываем файл эксель
# 8. Вставляем урл
# 9. Вставляем title
# 10. Вставляем description
# 11. Вставляем keywords

from urllib.request import urlopen
from bs4 import BeautifulSoup
import xlrd3 as xlrd
import xlwt

def get_urls_from_xlsx(count_url=100):
    urls = []
    try:
        workbook = xlrd.open_workbook('urls.xlsx')
        sheet = workbook.sheet_by_index(0)
        for rx in range(sheet.nrows):
            if rx == count_url:
                break
            row = sheet.row(rx)
            url = row[0].value
            urls.append(url)
    except Exception:
        print('Ошибка при открытии файла!\n\nПроверьте, название файла.\nПроверьте, что он находится в той же директории.\n')
    return urls

def get_url_content(url):
    html = urlopen(url).read().decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    title = soup.title.string if soup.title else "No title tag given"
    description = soup.find("meta", attrs={'name': 'description'})
    description = description["content"] if description else "No meta description given"
    keywords = soup.find("meta", attrs={'name': 'keywords'})
    keywords = keywords["content"] if keywords else "No meta keywords given"

    return (url, title, description, keywords)

def write_to_xls(urls, file_name='output.xls', count=1):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('1')
    sheet.write(0, 0, 'url')
    sheet.write(0, 1, 'title')
    sheet.write(0, 2, 'description')
    sheet.write(0, 3, 'keywords')
    workbook.save(file_name)
    for url in urls:
        url, title, description, keywords = get_url_content(url)
        print(url, title, description, keywords)
        sheet.write(count, 0, url)
        sheet.write(count, 1, title)
        sheet.write(count, 2, description)
        sheet.write(count, 3, keywords)
        count += 1
        workbook.save(file_name)
    print("\n\nurl, title, description, keywords записаны в файл output.xls")

if __name__ == '__main__':
    urls = get_urls_from_xlsx()
    write_to_xls(urls)



