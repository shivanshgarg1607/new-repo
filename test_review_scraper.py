from review_scraper import scrape_reviews

url = "https://www.hexahealth.com/delhi/hospital/apollo-hospital-delhi/reviews"

reviews = scrape_reviews(url)

print("=" * 60)
print(f"Found {len(reviews)} reviews\n")

for i, review in enumerate(reviews, start=1):
    print(f"Review {i}")
    print("-" * 40)
    print("Reviewer :", review["reviewer"])
    print("Rating   :", review["rating"])
    print("Date     :", review["date"])
    print("Review   :", review["review"])
    print()