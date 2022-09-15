from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Movies IMDB '
print(excel.sheetnames)
sheet.append(['Rank', 'Name', 'Year of Release','Certificate', 'Runtime', 'Genre1', 'Gnere2', 'Genre3', 'IMDB Rating', 'Metascore', 'Votes', 'Gross'])

'''
try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status() #if website is not reachable, gives an error

    soup = BeautifulSoup(source.text, 'html.parser')
    movies = soup.find('tbody', class_="lister-list").find_all('tr')

    for movie in movies:
        name = movie.find('td', class_='titleColumn').a.text
        rank = movie.find('td', class_='titleColumn').get_text(strip=True).split('.')[0]
        year = movie.find('td', class_="titleColumn").span.text.strip('()') #prints year, removes the ()
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text

        print(rank, name, year, rating)


except Exception as e:
   print(e) '''

try:
    source = requests.get('https://www.imdb.com/list/ls006266261/?sort=user_rating,desc&st_dt=&mode=detail&page=1')
    source.raise_for_status()

    soup= BeautifulSoup(source.text, 'html.parser')
    movies = soup.find(class_="lister-list").find_all(class_='lister-item-content')

    for movie in movies:
        rank = movie.find('span', class_='lister-item-index unbold text-primary').get_text(strip=True).split('.')[0]
        name = movie.find('h3', class_='lister-item-header').a.text
        year = movie.find('span', class_="lister-item-year text-muted unbold").text.strip('()')# prints year, removes the ()
        certificate = movie.find('span', class_="certificate").text
        runtime = movie.find('span', class_="runtime").text.replace('min', '')
        genre = movie.find('span', class_="genre").text.split(',')
        genre1 = genre[0]
        genre2= genre[1]  if len(genre)>1 else "null"
        genre3 = genre[2] if len(genre) > 2 else "null"
        rating = movie.find('span', class_="ipl-rating-star__rating").text
        metascore = movie.find('span', class_='metascore').text.replace(' ', '') if movie.find('span', class_="metascore") else 'null'
        #director = movie.find('p', class_='text-muted text-small') if movie.find('span', class_="text-muted text-small") else 'null'
        value = movie.find_all('span', attrs={'name':'nv'})
        votes = value[0].text
        gross = value[1].text.replace('$','').replace('M','') if len(value) > 1 else 'null'
        print(rank, name, year, certificate, runtime, genre1, genre2, genre3, rating, metascore, votes, gross )
        sheet.append([rank, name, year, certificate, runtime, genre1, genre2, genre3, rating, metascore, votes, gross])


except Exception as e:
    print(e)

excel.save('IMDB Movie Ratings.xlsx')



