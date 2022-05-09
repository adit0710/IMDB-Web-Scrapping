from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook() #to create new excel file

sheet =excel.active #storing excel sheet which is active in the variable called sheet
sheet.title = 'Top Rated Movie from IMDB' #giving name to the sheet
sheet.append(['RANK','MOVIE_NAME','RELEASE_DATE','RATING']) #added a row which is heading row in the sheet

#used because if it throws error then whole code will get crash
try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status() # method used to throw error if the url is wrong

    soup = BeautifulSoup(source.text,'html.parser')
     # beautifulsoup will take data in the text form by parsing in html page and store in soup variable
    #print(soup)

    movies = soup.find('tbody',class_="lister-list")
    #find method will find first tag with the given class value and assign in movie variable
    
    #print(movies) # will fetch whole table
    
    movies1 =movies.find_all('tr')
    #print(len(movies1)) #will fetch each entry in table

    for movie in movies1:

        rank = movie.find('td',class_="titleColumn").get_text(strip=True).split('.')[0] 
        #get_text to get the data and strip = true for ignoring the new line and
        #white space and split to get just the number as separate in output
        #and index 0 to get first item that is rank

        name = movie.find('td',class_="titleColumn").a.text 
        #will give you text which inside anchor tag so that is a

       

        year = movie.find('td',class_="titleColumn").span.text.strip('()')
        #it will take text from span tag and strip will remove the bracket from year

        rating = movie.find('td',class_="ratingColumn imdbRating").strong.text
        #same as name but instead of using anchor tag we are using strong tag as it is in website



        print(rank,name,year,rating) 
        sheet.append([rank,name,year,rating])
        
except Exception as e:
    print(e)

excel.save('IMDB MOVIE RATING.xlsx')

