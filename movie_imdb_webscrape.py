from bs4 import BeautifulSoup
import requests,openpyxl
from typing_extensions import runtime


excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies - IMDB'
sheet.append(['Movie Rank','Movie Name','Year of Release','Rating_IMDB'])
session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0..."})
try:
  source = session.get('https://www.imdb.com/chart/top/')
  source.raise_for_status()

  soup = BeautifulSoup(source.text,'html.parser')
  movies = soup.find('ul',class_="ipc-metadata-list ipc-metadata-list--dividers-between sc-9d2f6de0-0 iMNUXk compact-list-view ipc-metadata-list--base")
  for movie in movies:
    l1= movie.find('div', class_="ipc-metadata-list-summary-item__c")
    tl1 = l1.find('div',class_="ipc-metadata-list-summary-item__tc")
    mm=l1.find('h3',class_='ipc-title__text').text
    rank,name = mm.split('.')
    y1 = movie.find('div', class_="sc-479faa3c-7 jXgjdT cli-title-metadata")
    year = y1.find('span',class_="sc-479faa3c-8 bNrEFi cli-title-metadata-item").text
    r1 = movie.find('span',class_="sc-479faa3c-1 iMRvgp")
    rating = r1.find('span',class_="ipc-rating-star ipc-rating-star--base ipc-rating-star--imdb ratingGroup--imdb-rating").text.split('(')[0]
    
    sheet.append([rank,name,year,rating])

except Exception as e:
  print(e)

excel.save('IMDB_Movie_Ratings.xlsx')