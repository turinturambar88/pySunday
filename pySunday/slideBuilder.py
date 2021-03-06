
#Standard Library Imports
import os

#Anaconda Imports

#Local Imports
import hymn
import passage
import ppt


#Empty slides in Template
templates = {'black': 1, 'red': 2 }

hymn_title_format = ppt.TextboxFormat(
    left = 0, top = 0, width = 10, height = 1.25, 
    font_name = 'Gill Sans MT', alignment = ppt.alignCenter, font_size = 40,
    color = ppt.RGB(255,255,255), italic = True
)
hymn_body_format = ppt.TextboxFormat(
    left = 0.15, top = 1.5, width = 9.7, height = 6, 
    font_name = 'Gill Sans MT', alignment = ppt.alignCenter, font_size = 36,
    color = ppt.RGB(255,255,255), bold = True
)
scripture_title_format = ppt.TextboxFormat(
    left = 0.5, top = 0.35, width = 9, height = 1, 
    font_name = 'Arial Unicode MS', alignment = ppt.alignCenter, font_size = 44,
    color = ppt.RGB(255,255,204), shadow = True
)
scripture_body_format = ppt.TextboxFormat(
    left = 0.5, top = 1.25, width = 9, height = 5.9, 
    font_name = 'Tahoma',alignment = ppt.alignJustify, font_size = 32, 
    color = ppt.RGB(255,255,255), shadow = True
)


class SundaySlides:
    """
    """
    def __init__(self, filename, template = 'Template.pptx'):
        """
        """
        self.filename = filename
        self.template = template        
        self.esv_api = passage.ESV()
        self.powerpoint = ppt.PPTPres(fname = self.filename, template = self.template, visible = False)
        #blank slide at front        
        self._blank_slide()
    
    def _blank_slide(self):
        self.powerpoint.add_slide(master = templates['black'])
    
    def add_hymn(self, hymn_number):
        """
        Look up a hymn from http://www.hymnary.org/hymn/TH1990/ and create
        new PowerPoint slides.
        
        Not all hymns are available, some are missing verses, and hymns
        that include a refrain will work poorly.
        
        This method should only be used if a PowerPoint version of the hymn
        does not already exist.  Use "add_song" method if it does exist.
        """        
        my_hymn = hymn.Hymn(hymn_number) 
        my_hymn.scrape()
        my_hymn.split_verses()
        for i,verse in enumerate(my_hymn.verses):
            self.powerpoint.add_slide(master = templates['black'])
            slide_num = self.powerpoint.pres.Slides.Count
            if i == 0:
                self.powerpoint.add_textbox(slide_num, my_hymn.title, hymn_title_format)
            self.powerpoint.add_textbox(slide_num, verse, hymn_body_format)
        #blank slide after hymn
        self._blank_slide()

    def add_scripture(self, reference):
        """
        Look up scripture from http://www.esvapi.org/ and create new PowerPoint
        slides.
        
        You will likely have to clean up slides as the text will overflow in 
        some cases, or leave slides fairly empty in others.
        """
        self.esv_api.get_text_passage(reference)
        for block in self.esv_api.blocks:
            self.powerpoint.add_slide(master = templates['red'])
            slide_num = self.powerpoint.pres.Slides.Count
            self.powerpoint.add_textbox(slide_num, reference, scripture_title_format)
            self.powerpoint.add_textbox(slide_num, block, scripture_body_format)
        #blank slide after scripture        
        self._blank_slide()
    
    def add_song(self, filename):
        """
        Insert an existing powerpoint song.  
        
        This should be the preferred way to add songs 
        (once complete...NOT YET IMPLEMENTED)
        """
        if os.path.isfile(filename):
            pass            
            #Open file to copy from
            #Select all slides & copy
            #Paste into existing presentation
        else:
            print "File not found to add song: " + filename
            #Add placeholder slide and note that song is missing                    
            self.powerpoint.add_slide(master = templates['black'])
            slide_num = self.powerpoint.pres.Slides.Count
            self.powerpoint.add_textbox(
                slide_num, 
                "File not found: " + filename, 
                hymn_title_format
            )
    
    def save(self):
        """
        """
        
        #Delete template slides from front of presentation
        self.powerpoint.delete_slide(1)        
        self.powerpoint.delete_slide(1)    
        #Save presentation
        self.powerpoint.save()
        self.powerpoint.close()