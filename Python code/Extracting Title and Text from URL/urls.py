import requests
from bs4 import BeautifulSoup
import pandas as pd

path = "C:/Users/imran/Downloads/Input.xlsx"
df = pd.read_excel(path)
url_id = df['URL_ID'].to_list()
urls = df['URL'].to_list()


for id, links in zip(url_id, urls):
    data = requests.get(links)
    if (data.status_code == 200):
        print('Success', id)

        soup = BeautifulSoup(data.content, "html.parser")

        # lets extract 'Title' and 'Text'
        # Title
        title = soup.find('h1').text

        # Text
        content_div = soup.find('div', class_='td-post-content')    # class
        content = content_div.find_all('p')                         # tag

        # create empty string to hold the 'Text' Data located at tag 'p'
        extracted_data = ''
        for para in content:
            extracted_data += para.get_text() + '\n'

        # lets create a DataFrame to hold both 'Title' and 'Text'
        file = [[title, extracted_data]]
        pandasDF = pd.DataFrame(file, columns = ['Title', 'Text'])

        # create a csv file
        CSV_file = pandasDF.to_excel(f"{id}.xlsx", index = False)
    else:
        print("Error ->", data, " ", id)
