from os import execl
import requests 
from bs4 import BeautifulSoup
import openpyxl 


excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Top Rated Movies"
sheet.append(['Rank','Title','Year','Rating'])




url  = requests.get('https://www.imdb.com/chart/top/')
result = url.content

soup = BeautifulSoup(result,'lxml')
movies = soup.find('tbody',{"class":"lister-list" }).find_all('tr')


for movie in movies : 
    title = movie.find('td',class_='titleColumn').find('a').text
    rank =  movie.find('td',class_='titleColumn').getText(strip=True).split(".")[0]
    year = movie.find('td',class_='titleColumn').span.text.strip('()')
    rating = movie.find('td',class_="ratingColumn imdbRating").strong.text

    sheet.append([rank,title,year,rating])

excel.save('Top Rating movies.xlsx')

