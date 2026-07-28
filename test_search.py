import requests
from bs4 import BeautifulSoup
from urllib.parse import quote

hospital = "Yashoda Super Speciality Hospital"

query = quote(f"{hospital} reviews")

url = f"https://html.duckduckgo.com/html/?q={query}"

headers = {
    "User-Agent": "Mozilla/5.0"
}

print("Searching...")
print(url)

response = requests.get(url, headers=headers)

print("Status:", response.status_code)

soup = BeautifulSoup(response.text, "html.parser")

results = soup.select(".result")

print(f"\nFound {len(results)} results\n")

for i, result in enumerate(results[:10], start=1):

    title = result.select_one(".result__title")

    link = result.select_one(".result__url")

    if title:
        print("=" * 60)
        print(i)
        print(title.get_text(" ", strip=True))

        if link:
            print(link.get_text(" ", strip=True))