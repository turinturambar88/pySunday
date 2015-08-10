
#Standard Library
import os

#Local
from pySunday import ppt


class CCLIFixer():
    def __init__(self, original_fname, final_fname):
        self.original_fname = original_fname
        self.final_fname = final_fname
        
        #Create a copy to work with
        os.copyfile(self.original_fname, self.final_fname)
        
        self.parse_text()
        self.get_ccli_info()
        self.apply_ccli_info()
        
    def parse_text():
        """
        Pull song text from powerpoint
        """
    
    def get_ccli_info(self):
        """
        Find some API or website to scrape info based on song text
        Use "requests" module
        """
    
    def apply_ccli_info(self):
        """
        Add CCLI into text box to powerpoint
        """
