from googlesearch import search
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse
import time

courses = []

def has_any_class(tag):
    return tag.name == "tr" and tag.has_attr("class") and any(c in tag["class"] for c in class_names)
class_names = ["tblrowAlt", "tblrow"]  # Add your class names here

def search_and_download_pdf():

        response = requests.get("https://ocw.vu.edu.pk/", verify=False)
        soup = BeautifulSoup(response.text, 'html.parser')
        links = soup.find_all('a', href=True)
        for link in links:
            href = link['href']
            if "Courses.aspx?cat=" in href:
                catUrl = "https://ocw.vu.edu.pk/" +  href
                res = requests.get(catUrl, verify=False)
                soup = BeautifulSoup(res.text, 'html.parser')
                rows = soup.find_all(has_any_class)
                for row in rows:
                    cells = row.find_all('td')
                    fstr = ""
                    for cell in cells:
                        anchor = cell.find('a')
                        if anchor:
                             fstr = fstr + " " + anchor.string.strip()
                    courses.append(fstr)
        return

search_and_download_pdf()
f = open('vu_all_courses_names.txt', 'w')
for course in courses:
    f.write(course + '\n')
f.close()
print("File generated Successfully")