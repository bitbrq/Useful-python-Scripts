import requests
from bs4 import BeautifulSoup


def search_google_maps(location, keyword):
    url = f"https://www.google.com/maps/@32.5779456,71.5358208,13z?entry=ttu/{keyword}+in+{location}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"}

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()  # Raise exception for 4xx and 5xx status codes
        return response.content
    except requests.RequestException as e:
        print("Failed to fetch data from Google Maps:", e)
        return None


def parse_google_maps(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    places = soup.find_all("div", class_="section-result-text")

    parsed_data = []
    for place in places:
        name = place.find("h3", class_="section-result-title").text.strip()
        address = place.find("span", class_="section-result-location").text.strip()
        rating_tag = place.find("span", class_="cards-rating-score")
        rating = rating_tag.text.strip() if rating_tag else "N/A"
        phone_tag = place.find("span", class_="section-result-phone-number")
        phone = phone_tag.text.strip() if phone_tag else "N/A"
        parsed_data.append({"Name": name, "Address": address, "Rating": rating, "Phone": phone})

    return parsed_data


def save_to_file(data, filename):
    with open(filename, 'w', encoding='utf-8') as f:
        for item in data:
            f.write(f"Name: {item['Name']}\n")
            f.write(f"Address: {item['Address']}\n")
            f.write(f"Phone: {item['Phone']}\n")
            f.write(f"Rating: {item['Rating']}\n\n")


def main():
    location = input("Enter the location: ")
    keyword = input("Enter the keyword: ")

    html_content = search_google_maps(location, keyword)
    if html_content:
        parsed_data = parse_google_maps(html_content)
        if parsed_data:
            filename = f"{location}_{keyword}_results.txt"
            save_to_file(parsed_data, filename)
            print(f"Data saved to {filename}")
        else:
            print("No search results found.")
    else:
        print("No data to save.")


if __name__ == "__main__":
    main()
