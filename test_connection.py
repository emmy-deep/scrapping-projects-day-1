
import requests

url = "https://books.toscrape.com/"

try:
    response = requests.get(url, timeout=15)

    print("Status:", response.status_code)
    print("Website:", response.url)

    response.raise_for_status()

    print("Connection successful!")

except requests.exceptions.ProxyError as e:
    print("Proxy problem:", e)

except requests.exceptions.RequestException as e:
    print("Request problem:", e)
