
import requests
from bs4 import BeautifulSoup
import pandas as pd

url = "https://books.toscrape.com/"

headers = {
    "User-Agent": "MyLearningScraper/1.0"
}

response = requests.get(
    url,
    headers=headers,
    timeout=15
)

response.raise_for_status()

soup = BeautifulSoup(response.text, "html.parser")

books = soup.select("article.product_pod")

data = []

for book in books:

    # Extract book information
    title = book.h3.a["title"]

    price = book.select_one(
        ".price_color"
    ).get_text(strip=True)

    availability = book.select_one(
        ".availability"
    ).get_text(strip=True)

    rating = book.select_one(
        "p.star-rating"
    )["class"][1]

    # Store information without website links
    data.append({
        "Book Name": title,
        "Price": price,
        "Availability": availability,
        "Rating": rating
    })

# Convert extracted data into Excel format
df = pd.DataFrame(data)

# Save Excel file
df.to_excel(
    "books_data.xlsx",
    sheet_name="Books",
    index=False
)

print("Scraping completed!")
print("Total records:", len(data))
print("Excel file created: books_data.xlsx")
