import requests
from bs4 import BeautifulSoup
from xlwt import Workbook


# Create workbook
wb = Workbook()
sheet = wb.add_sheet('Sheet 1') # add_sheet is used to create sheet.

# Define column name
sheet.write(0, 0, 'MOVIE_NAME')
sheet.write(0, 1, 'RELEASE_YEAR')
sheet.write(0, 2, 'RATINGS')

url = "https://www.imdb.com/chart/top/"

response = requests.get(url)
html = response.content
soup = BeautifulSoup(html,'html.parser')
#print(soup)
movies = soup.find('tbody',class_="lister-list").find_all('tr')

n=1
for movie in movies:
    name = movie.find('td',class_="titleColumn").a.text
    release_date = movie.find('td',class_="titleColumn").span.text
    rating = movie.find('td',class_="ratingColumn imdbRating").strong.text

    sheet.write(n,0,name)
    sheet.write(n,1,release_date)
    sheet.write(n,2,rating)
    n = n+1

wb.save('Imdb.xls')
    
