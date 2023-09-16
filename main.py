import requests, openpyxl
from bs4 import BeautifulSoup

excel = openpyxl.Workbook()
print(excel.sheetnames)

sheet = excel.active
sheet.title = "Top rated movies"
print(excel.sheetnames)

sheet.append(["Movie rRank", "Movie Name", "Year of Release", "IMDB Rating"])

def fetch_data_and_load_to_excel():
    try:
        website_url = "https://www.imdb.com/chart/top/"
        user_agent = {'User-agent': 'Mozilla/5.0'}

        # Making a get request to fetch the dat
        response = requests.get(website_url, headers=user_agent)

        # This will throw the error , if some issue occurs with the request
        response.raise_for_status()

        # Converting to the beautiful soup form
        soup = BeautifulSoup(response.text, "html.parser")

        movies = soup.find("ul", class_="ipc-metadata-list").find_all("li")

        for movie in movies:
            tag = movie.find("div", class_="ipc-metadata-list-summary-item__tc")

            rank = tag.a.text.split(".")[0]
            name = tag.a.text.split(".")[1].strip("")
            year = tag.find("div", class_="cli-title-metadata").find_all("span")[0].text
            rating = movie.find("span", class_="ipc-rating-star").text.split("(")[0]

            print(rank, name, year, rating)

            sheet.append([rank, name, year, rating])

    except Exception as e:
        print(e)


fetch_data_and_load_to_excel()

excel.save("IMDB Movie Ratings.xlsx")
