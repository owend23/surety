import requests
from bs4 import BeautifulSoup

def get_soup(url):
	r = requests.get(url)
	return BeautifulSoup(res.text, 'html.parser')

