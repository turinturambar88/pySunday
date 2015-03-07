
#Standard Library Imports
import re

#Anaconda Imports
from bs4 import BeautifulSoup
import requests

#Local Imports


class Hymn:
    """
    """
    def __init__(self, number, base_url = 'http://www.hymnary.org/hymn/TH1990/'):
        self.number = str(number)
        self.base_url = base_url
        self.title = ''
        self.text = ''
        self.verses = []
    
    def scrape(self):
        full_page = requests.get(self.base_url + self.number)
        if full_page.ok is True:
            soup = BeautifulSoup(full_page.text)
            hymn_page = soup.find('div', {'class': 'hymnpage'})
            
            hymn_title = hymn_page.find('h2', {'class':'hymntitle'})
            self.title = hymn_title.text
            self.title = re.sub(self.number+'. ','',self.title)
            self.title += '\nHymn #' + self.number

            hymn_text = hymn_page.find('div', {'id':'text'})
            if hymn_text is not None:
                self.text = hymn_text.text
            else:
                self.text = "Text not available online"
        else:
            print "Error scraping Hymn #" + str(self.number)
        
    def split_verses(self):
        """
        Break hymn text into verses
        """
        self.verses = re.split('[0-9]+',self.text)
        self.verses.pop(0) # strip empty entry


if __name__ == '__main__':
    my_hymn = Hymn(230) 
    my_hymn.scrape()
    my_hymn.split_verses()
    
#    site = requests.get('http://www.hymnary.org/hymn/TH1990/100')
#    if site.ok is True:
#        soup = BeautifulSoup(site.text)
#        hymn_page = soup.find('div', {'class': 'hymnpage'})
#        hymn_title = hymn_page.find('h2', {'class':'hymntitle'})
#        print hymn_title.text
#        hymn_text = hymn_page.find('div', {'id':'text'})
#        print hymn_text.text
#    else:
#        print "Error loading Hymn"