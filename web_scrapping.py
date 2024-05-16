import requests
from bs4 import BeautifulSoup
import xlwt
import warnings
from urllib3.exceptions import InsecureRequestWarning

# Suppress only the single InsecureRequestWarning from urllib3 needed to silence the warning.
warnings.simplefilter('ignore', InsecureRequestWarning)

def scrape_data_and_generate_excel(base_url):
    response = requests.get(base_url, verify=False)
    soup = BeautifulSoup(response.content, "html.parser")

    # Determine the number of pages
    pagination = soup.find("ul", class_="page-list clearfix text-xs-right")
    last_page = int(pagination.find_all("li")[-2].text) if pagination else 1

    data = []
    for page_number in range(1, last_page + 1):
        url = f"{base_url}?page={page_number}" if page_number > 1 else base_url
        response = requests.get(url, verify=False)
        soup = BeautifulSoup(response.content, "html.parser")

        product_titles = [elem.text.strip() for elem in soup.find_all("div", class_="h3 product-title")]
        product_marques = [elem.text.strip() for elem in soup.find_all("div", class_="div_manufacturer_name")]
        product_descriptions = [elem.text.strip() for elem in soup.find_all("div", class_="div_contenance")]
        original_prices = [elem.text.strip() for elem in soup.find_all("div", class_="regular-price")]
        final_prices = [elem.text.strip() for elem in soup.find_all("span", class_="price")]

        for i in range(len(product_titles)):
            data.append({
                "name": product_titles[i] if i < len(product_titles) else "",
                "marque": product_marques[i] if i < len(product_marques) else "",
                "description": product_descriptions[i] if i < len(product_descriptions) else "",
                "original_price": original_prices[i] if i < len(original_prices) else "",
                "final_price": final_prices[i] if i < len(final_prices) else "",
            })

        if last_page == 1:
            break

    file_name = "monoprix_boissons.xls"
    write_data_to_excel(data, file_name)
    return data, file_name

def write_data_to_excel(data, file_name):
    workbook = xlwt.Workbook(encoding="utf-8")
    sheet = workbook.add_sheet("Products")

    headers = ["name", "marque", "description", "original_price", "final_price"]
    for col, header in enumerate(headers):
        sheet.write(0, col, header)

    for row, item in enumerate(data, start=1):
        sheet.write(row, 0, item["name"])
        sheet.write(row, 1, item["marque"])
        sheet.write(row, 2, item["description"])
        sheet.write(row, 3, item["original_price"])
        sheet.write(row, 4, item["final_price"])

    workbook.save(file_name)

if __name__ == "__main__":
    url_to_scrape = "https://courses.monoprix.tn/13-boissons"
    data, file_name = scrape_data_and_generate_excel(url_to_scrape)
    print(f"Data has been scraped and saved to {file_name}")
