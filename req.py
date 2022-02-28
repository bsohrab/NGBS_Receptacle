import requests
from bs4 import BeautifulSoup
import urllib.request
# Create an OpenerDirector with support for Basic HTTP Authentication...
auth_handler = urllib.request.HTTPBasicAuthHandler()
auth_handler.add_password(realm='PDQ Application',
                      uri='https://mahler:8092/site-updates.py',
                      user='klem',
                      passwd='kadidd!ehopper')

opener = urllib.request.build_opener(auth_handler)

# ...and install it globally so it can be used with urlopen.
urllib.request.install_opener(opener)
f = urllib.request.urlopen('http://www.example.com/login.html')
csv_content = f.read()
page = requests.get("https://axis.pivotalenergy.net/home/139570/#/tabs/qa")
soup = BeautifulSoup(page.content, 'html.parser')
body = list(soup.children)
print( soup.find_all('a'))# class_="ng-binding"))

