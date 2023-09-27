import csv
import json
from collections import OrderedDict
import shutil
import pendulum
import pandas as pd
import requests
import time
from tempfile import NamedTemporaryFile
from bs4 import BeautifulSoup

def read_excel_file(is_csv):
    all_urls = []
    file_path = "data.xlsx"
    xls = pd.ExcelFile(file_path)

    sheetX = xls.parse(0)

    urls = sheetX['LINK']
    keywords = sheetX['Keyword']
    
    for url, keyword in zip(urls, keywords):
        try:
            if 'http' in url:
                url = f"{url}?keyword={keyword}"
                all_urls.append(url)
        except:
            pass
    
    if is_csv == 0:
        try:
            with open('cejene_data.csv', 'a') as f:
                writer = csv.writer(f)
                writer.writerow(["Link", "Keyword", "Status", "Mesto", "Nasa-Cena", "Cena-nad-nami", "Cena-pod-nami"])
        except:
            pass
    
    return all_urls


def read_csv_data():
    with open('cejene_data.csv') as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        for row in readCSV:
            if len(row) > 0:
                url = f"{row[0]}?keyword={row[1]}"
                try:
                    if 'http' in row[0]:
                        time.sleep(1)
                        extract_data(url, False, row)
                except:
                    extract_data(url, True, row)

def extract_data(url, use_proxy, row_data=None):
    try:
        url_split = url.split('?keyword=')
        main_url = url_split[0]
        keyword_to_search = url_split[1]
        print("Fetching html against url {0}".format(main_url))
        
        page = requests.get(main_url)
        soup = BeautifulSoup(page.content, 'html.parser')
                
        main_status = "Not-Available"
        above_price = None
        below_price = None
        main_price = None

        all_keys = [item['alt'] for item in soup.css.select("div#offersList div.topRow div.logotip a > img")]

        all_values = [item.text.split('\n')[3].strip() for item in soup.css.select('div#offersList div.priceBox div.price a')]

        d = OrderedDict(zip(all_keys, all_values))
        d = OrderedDict(d)
        d = json.dumps(d)
        d = json.loads(d)
        
        main_index = None
        print("Product to search is {0}".format(keyword_to_search))
        
        if keyword_to_search in all_keys:
            print("Product Found")
            main_status = "Available"
            main_index = all_keys.index(keyword_to_search)
            main_price = d.get(keyword_to_search)

            prev_index = main_index - 1
            next_index = main_index + 1
            main_index = main_index + 1
            prev_index_key = all_keys[prev_index]
            if prev_index_key:
                above_price = d.get(prev_index_key)
            next_index_key = all_keys[next_index]
            if next_index_key:
                below_price = d.get(next_index_key)
        else:
            print("Product Not Available")
        
        if not row_data:
            with open('cejene_data.csv', 'a') as f:
                writer = csv.writer(f)
                writer.writerow([main_url, keyword_to_search, main_status, main_index, main_price, above_price, below_price])
            f.close()
        else:
            main_data = [main_url,keyword_to_search]
            past_data = row_data[6:10]
            current_data = [main_status, main_index, main_price, pendulum.now()]
            update_data(main_data+past_data+current_data)
    except:
        pass

def update_data(row_data):
    filename = 'cejene_data.csv'
    tempfile = NamedTemporaryFile(mode='w', delete=False)

    fields = ["Link", "Keyword", "Status", "Mesto",
              "Nasa-cena", "Cena-nad-nami", "cena-pod-nami"]

    with open(filename, 'r') as csvfile, tempfile:
        reader = csv.DictReader(csvfile, fieldnames=fields)
        writer = csv.DictWriter(tempfile, fieldnames=fields)
        
        for row in reader:
            if row['Link'] == row_data[0]:
                row = {'Link': row_data[0], 'Keyword': row_data[1],
                       'Status': row_data[2], 'Mesto': row_data[3],
                       'Nasa-cena': row_data[4], 'Cena-nad-nami': row_data[5], 'Cena-pod-nami': row_data[6]
                       }
            writer.writerow(row)

    shutil.move(tempfile.name, filename)


def main_function():
    file_exist = False
    csv_data = 0
    
    try:
        with open('cejene_data.csv') as f:
            readCSV = csv.reader(f, delimiter=',')
            csv_data = len(list(readCSV))
            file_exist = True
    except IOError:
        pass
        
    if file_exist and csv_data > 100:
        read_csv_data()
    else:
        urls = read_excel_file(csv_data)
        for url in urls:
            try:
                extract_data(url, False, None)
                time.sleep(3)
            except Exception as ex:
                extract_data(url, True, None)

if __name__ == '__main__':
    main_function()
