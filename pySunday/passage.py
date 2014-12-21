#!/usr/bin/env python

#Standard Library Imports
import urllib

#Anaconda Imports

#Local Imports


class ESV:
    def __init__(self, line_length = 40):
        self.base_url = 'http://www.esvapi.org/v2/rest/verse?key=IP'
        options = [
            'include-footnotes=0',
            'include-short-copyright=0',
            'output-format=plain-text',
            'include-passage-horizontal-lines=0',
            'include-heading-horizontal-lines=0',
            'include-headings=0',
            'include-subheadings=0',
        ]
        self.options = '&'.join(options)
        self.options += 'line-length=' + str(line_length)
        self.result = ''

    def get_text_passage(self, passage):
        passage = '+'.join(passage.split())
        url = self.base_url + '&passage=' + passage + '&' + self.options
        page = urllib.urlopen(url)
        self.result = page.read()
        
        lines = self.result.split('\r\n')
        blocks = []
        #11 lines max
        
        
        


if __name__ == '__main__':

    esv_api = ESV()
    print esv_api.get_text_passage("John 3-5:3")

