#!/usr/bin/env python

#Standard Library Imports
import urllib
import re

#Anaconda Imports

#Local Imports


class ESV:
    def __init__(self, line_length = 0):
        self.base_url = 'http://www.esvapi.org/v2/rest/verse?key=IP'
        options = [
            'include-passage-references=0',
            'include-first-verse-numbers=1',
            'include-footnotes=0',
            'include-short-copyright=0',
            'output-format=plain-text',
            'include-passage-horizontal-lines=0',
            'include-heading-horizontal-lines=0',
            'include-headings=0',
            'include-subheadings=0',
        ]
        self.options = '&'.join(options)
        self.options += '&line-length=' + str(line_length)
        self.result = ''

    def get_text_passage(self, passage):
        passage = '+'.join(passage.split())
        url = self.base_url + '&passage=' + passage + '&' + self.options
        print url
        page = urllib.urlopen(url)
        self.result = page.read()
        
        #Format 
        self.result = self.result.replace('[','')
        self.result = self.result.replace(']',' ')
        self.result = self.result.replace('\n\n','\n')
        
        self.split_into_blocks()

    def split_into_blocks(self):        
        """
        Split into blocks to fill individual powerpoint slides
        """
        self.words = self.result.split(' ')
        
        self.blocks = []
        block_start = 0
        block_end = 0
        block_length = 0
        
        screen_length_max = 475 #max characters on a screen
        screen_length_buffer = 75 #allow stopping this many characters early if a sentence ends        
        
        for word in self.words:
            block_end += 1
            block_length += len(word) + 1
            if '\n' in word:
                #Special case...poetry or end of paragraph...account for extra space used
                block_length += 20                
            end_of_sentence = (re.search('[\?\!\.\;\"\']',word) is not None)

            inside_buffer = block_length >= (screen_length_max - screen_length_buffer)

            if (end_of_sentence and inside_buffer) or (block_length >= screen_length_max):            
                self.blocks.append(' '.join(self.words[block_start:block_end]))
                block_start = block_end
                block_length = 0
        else:
            #Final screen
            self.blocks.append(' '.join(self.words[block_start:]))



if __name__ == '__main__':

    esv_api = ESV()
    print esv_api.get_text_passage("John 3-5:3")

    esv_api2 = ESV()
    esv_api2.get_text_passage("James 3:13-17")
    print esv_api2.result    
    print esv_api2.blocks
