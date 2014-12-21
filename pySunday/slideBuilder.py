
#Standard Library Imports


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
    left = 0.15, top = 1.25, width = 9.7, height = 6, 
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
    font_name = 'Tahoma',alignment = ppt.alignCenter, font_size = 32, 
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
        """
        self.esv_api.get_text_passage(reference)
        print self.esv_api.result 
        
        #blank slide after scripture        
        self._blank_slide()
    
    def save(self):
        """
        """
         
        
        
        #Delete template slides from front of presentation
        self.powerpoint.delete_slide(1)        
        self.powerpoint.delete_slide(1)    
        #Save presentation
        self.powerpoint.save()
        self.powerpoint.close()