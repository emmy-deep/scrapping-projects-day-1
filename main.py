
import requests
import pandas as pd
import os
import time

from bs4 import BeautifulSoup
from urllib.parse import urljoin
from datetime import datetime


# ==========================================
# CONFIGURATION
# ==========================================

OUTPUT_FOLDER = "output"
OUTPUT_FILE = os.path.join(
    OUTPUT_FOLDER,
    "scraping_results.xlsx"
)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; LearningScraper/1.0)"
}

session = requests.Session()
session.headers.update(HEADERS)


# ==========================================
# COMMON REQUEST FUNCTION
# ==========================================

def get_page(url):

    try:
        response = session.get(
            url,
            timeout=20
        )

        response.raise_for_status()

        return response.text

    except requests.RequestException as error:

        print(f"Request failed: {url}")
        print(error)

        return None


# ==========================================
# 1. BOOKS TO SCRAPE
# ==========================================

def scrape_books():

    print("\nStarting Books Scraping...")

    url = "https://books.toscrape.com/"

    books_data = []

    while url:

        html = get_page(url)

        if not html:
            break

        soup = BeautifulSoup(html, "html.parser")

        books = soup.select("article.product_pod")

        for book in books:

            title_tag = book.select_one("h3 a")
            price_tag = book.select_one(".price_color")
            stock_tag = book.select_one(".availability")
            rating_tag = book.select_one(".star-rating")

            title = (
                title_tag.get("title", "").strip()
                if title_tag else ""
            )

            price = (
                price_tag.get_text(strip=True)
                if price_tag else ""
            )

            availability = (
                stock_tag.get_text(" ", strip=True)
                if stock_tag else ""
            )

            rating = (
                rating_tag.get("class", [])[-1]
                if rating_tag else ""
            )

            product_link = (
                urljoin(url, title_tag.get("href"))
                if title_tag else ""
            )

            books_data.append({
                "Title": title,
                "Price": price,
                "Availability": availability,
                "Rating": rating,
                "Product Link": product_link
            })

        print(f"Books collected: {len(books_data)}")

        # Find next page

        next_button = soup.select_one("li.next a")

        if next_button:

            url = urljoin(
                url,
                next_button.get("href")
            )

            time.sleep(1)

        else:

            url = None

    print("Books scraping completed!")

    return books_data


# ==========================================
# 2. QUOTES TO SCRAPE
# ==========================================

def scrape_quotes():

    print("\nStarting Quotes Scraping...")

    url = "https://quotes.toscrape.com/"

    quotes_data = []

    while url:

        html = get_page(url)

        if not html:
            break

        soup = BeautifulSoup(html, "html.parser")

        quotes = soup.select(".quote")

        for quote in quotes:

            text_tag = quote.select_one(".text")
            author_tag = quote.select_one(".author")

            tags = quote.select(".tags a.tag")

            quote_text = (
                text_tag.get_text(strip=True)
                if text_tag else ""
            )

            author = (
                author_tag.get_text(strip=True)
                if author_tag else ""
            )

            tag_names = ", ".join(
                tag.get_text(strip=True)
                for tag in tags
            )

            quotes_data.append({
                "Quote": quote_text,
                "Author": author,
                "Tags": tag_names
            })

        print(f"Quotes collected: {len(quotes_data)}")

        # Pagination

        next_button = soup.select_one("li.next a")

        if next_button:

            url = urljoin(
                url,
                next_button.get("href")
            )

            time.sleep(1)

        else:

            url = None

    print("Quotes scraping completed!")

    return quotes_data


# ==========================================
# 3. FAKE STORE API
# ==========================================

def scrape_fake_store():

    print("\nStarting Fake Store API...")

    url = "https://fakestoreapi.com/products"

    try:

        response = session.get(url, timeout=20)

        response.raise_for_status()

        products = response.json()

        products_data = []

        for product in products:

            products_data.append({

                "Product ID": product.get("id"),

                "Title": product.get("title"),

                "Price": product.get("price"),

                "Category": product.get("category"),

                "Description": product.get("description"),

                "Image": product.get("image"),

                "Rating": product.get("rating", {}).get("rate"),

                "Review Count": product.get("rating", {}).get("count")

            })

        print(
            f"Products collected: {len(products_data)}"
        )

        return products_data

    except requests.RequestException as error:

        print("Fake Store API error:", error)

        return []

    except ValueError as error:

        print("Invalid API response:", error)

        return []


# ==========================================
# 4. JSONPLACEHOLDER POSTS
# ==========================================

def scrape_posts():

    print("\nStarting JSONPlaceholder...")

    url = "https://jsonplaceholder.typicode.com/posts"

    try:

        response = session.get(url, timeout=20)

        response.raise_for_status()

        posts = response.json()

        posts_data = []

        for post in posts:

            posts_data.append({

                "Post ID": post.get("id"),

                "User ID": post.get("userId"),

                "Title": post.get("title"),

                "Body": post.get("body")

            })

        print(
            f"Posts collected: {len(posts_data)}"
        )

        return posts_data

    except requests.RequestException as error:

        print("Posts API error:", error)

        return []

    except ValueError as error:

        print("Invalid posts response:", error)

        return []


# ==========================================
# 5. SAVE ALL DATA TO EXCEL
# ==========================================

def save_to_excel(
    books,
    quotes,
    products,
    posts
):

    os.makedirs(
        OUTPUT_FOLDER,
        exist_ok=True
    )

    data_sheets = {

        "Books": books,

        "Quotes": quotes,

        "Products": products,

        "Posts": posts

    }

    try:

        with pd.ExcelWriter(
            OUTPUT_FILE,
            engine="openpyxl"
        ) as writer:

            for sheet_name, data in data_sheets.items():

                df = pd.DataFrame(data)

                df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

                # Format Excel columns

                worksheet = writer.sheets[sheet_name]

                worksheet.freeze_panes = "A2"

                worksheet.auto_filter.ref = (
                    worksheet.dimensions
                )

                for column_cells in worksheet.columns:

                    max_length = max(
                        len(str(cell.value or ""))
                        for cell in column_cells
                    )

                    column_letter = (
                        column_cells[0].column_letter
                    )

                    worksheet.column_dimensions[
                        column_letter
                    ].width = min(max_length + 3, 60)

        print("\nExcel file created successfully!")

        print("File location:", os.path.abspath(OUTPUT_FILE))

    except Exception as error:

        print("Excel saving error:", error)


# ==========================================
# 6. MAIN EXECUTION
# ==========================================

def main():

    print("=" * 50)

    print("MULTI-WEBSITE WEB SCRAPING SYSTEM")

    print("=" * 50)

    start_time = datetime.now()

    books = scrape_books()

    quotes = scrape_quotes()

    products = scrape_fake_store()

    posts = scrape_posts()

    save_to_excel(
        books,
        quotes,
        products,
        posts
    )

    end_time = datetime.now()

    print("\n" + "=" * 50)

    print("SCRAPING SUMMARY")

    print("=" * 50)

    print("Total Books:", len(books))

    print("Total Quotes:", len(quotes))

    print("Total Products:", len(products))

    print("Total Posts:", len(posts))

    print("Start Time:", start_time)

    print("End Time:", end_time)

    print("Time Taken:", end_time - start_time)

    print("\nPROCESS COMPLETED!")


# ==========================================
# RUN PROGRAM
# ==========================================

if __name__ == "__main__":

    main()
