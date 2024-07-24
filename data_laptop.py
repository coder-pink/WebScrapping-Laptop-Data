import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

url = "https://www.amazon.com/s?k=laptop&ref=nb_sb_noss"

#Headers for requests
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3',
    'Accept-Language': 'en-US, en;q=0.5'
}

#http requests
response = requests.get(url, headers=HEADERS)
# print(response)
# print(type(response.content))

soup = BeautifulSoup(response.content, 'html.parser')
# print(soup)

# Find all the links with the specified class
links = soup.find_all("a", class_="a-link-normal s-underline-text s-underline-link-text s-link-style a-text-normal")

#store the link
links_list = []

# loop to extract tag objects from links
for link in links:
    links_list.append(link.get("href"))

# dictionary

product_details = {"title": [], "price": [], "rating": [], "reviews": []}


#loop to extract product details from link_list
for link in links_list:
    new_response = requests.get("https://www.amazon.com" + link, headers=HEADERS)

    new_soup = BeautifulSoup(new_response.content, "html.parser")


# get the title , price, rating and reviews


    # Get the title
    title = new_soup.find("span", {"id": "productTitle"})
    title_text = title.get_text(strip=True) if title else "N/A"

    # Get the price
    price = new_soup.find("span", {"class": "a-price-whole"})
    price_text = price.get_text(strip=True) if price else "N/A"

    # Get the rating
    rating = new_soup.find("span", {"class": "a-icon-alt"})
    rating_text = rating.get_text(strip=True) if rating else "N/A"

    # Get the number of reviews
    reviews = new_soup.find("span", {"id": "acrCustomerReviewText"})
    reviews_text = reviews.get_text(strip=True) if reviews else "N/A"

    # Append the extracted data to the dictionary
    product_details["title"].append(title_text)
    product_details["price"].append(price_text)
    product_details["rating"].append(rating_text)
    product_details["reviews"].append(reviews_text)

# Create an Excel workbook and sheet
wb = Workbook()
ws = wb.active
ws.title = "Amazon Laptop Data"


# Write the headers
headers = ["Title", "Price", "Rating", "Reviews"]
ws.append(headers)


# Write the data
for i in range(len(product_details["title"])):
    row = [
        product_details["title"][i],
        product_details["price"][i],
        product_details["rating"][i],
        product_details["reviews"][i]
    ]
    ws.append(row)

# Save the workbook
wb.save("Amazon_Laptop_Data.xlsx")

print("Data has been written to Amazon_Laptop_Data.xlsx")


