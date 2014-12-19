
#Standard Library Imports
import os

#Anaconda Imports
import win32com.client


#Local Imports

inches_to_points = 72.0


#NEED TO GET MORE POWERPOINT HARDCODES SET TO VARIABLES HERE!!

ppLayoutBlank = 12

#Text Alignment
alignLeft = 1 
alignCenter = 2
alignRight = 3

def RGB(red, green, blue):
    """
    :param red: Color value for red (0-255)
    :type red: int
    :param red: Color value for green (0-255)
    :type red: int
    :param red: Color value for blue (0-255)
    :type red: int
    
    Returns a color code for PowerPoint
    """
    return (blue << 16) | (green << 8) | (red)

class PictureFormat:
    """
    :param left: Location of left side of picture in inches from left side of slide
    :type left: float
    :param top: Location of top of picture in inches from top side of slide
    :type top: float
    :param width: Width of picture in inches
    :type width: float
    :param height: Height of picture in inches
    :type height: float
    :param crop_left: Amount of picture in inches to crop from left side.  Inches refer to picture's original size.
    :type crop_left: float
    :param crop_top: Amount of picture in inches to crop from top.  Inches refer to picture's original size.
    :type crop_top: float
    :param crop_right: Amount of picture in inches to crop from right side.  Inches refer to picture's original size.
    :type crop_right: float
    :param crop_bottom: Amount of picture in inches to crop from bottom.  Inches refer to picture's original size.
    :type crop_bottom: float
    
    If both width and height are None, the picture will be inserted at its original size
    
    Setting only height or width will keep original aspect ratio
    """
    def __init__(self, left, top, width = None, height = None, crop_left = 0, crop_top = 0, crop_right = 0, crop_bottom = 0):
        self.left = left * inches_to_points
        self.top = top * inches_to_points
        
        self.width = width
        if self.width is not None:
            self.width *= inches_to_points
        self.height = height
        if self.height is not None:
            self.height *= inches_to_points
            
        self.crop_left = crop_left * inches_to_points
        self.crop_top = crop_top * inches_to_points
        self.crop_right = crop_right  * inches_to_points
        self.crop_bottom = crop_bottom  * inches_to_points


class TextboxFormat:
    """
    
    """
    def __init__(self, left, top, width, height, font_name = 'Arial', font_size = 12, alignment = alignLeft, color = RGB(0, 0, 0), bold = False, italic = False, underline = False):
        self.left = left * inches_to_points
        self.top = top * inches_to_points
        self.width = width * inches_to_points
        self.height = height * inches_to_points
        self.font_name = font_name
        self.font_size = font_size
        self.alignment = alignment
        self.color = color
        self.bold = bold
        self.italic = italic
        self.underline = underline

class PPTPres:
    """
    :param fname: Filename of output PowerPoint presentation
    :type fname: string
    :param template: Filename of existing PowerPoint to load as a template
    :type template: string
    :param visible: Flag to open PowerPoint for viewing while performing actions.  Set to False for large batch operation
    :type visible: bool
    """
    def __init__(self, fname, template = None, visible = True):
        self.fname = os.path.abspath(fname)
        self.template = template
        
        #PowerPoint application
        self.ppt = win32com.client.Dispatch("PowerPoint.Application")
        
        if self.template is None:
            self.pres = self.ppt.Presentations.Add(WithWindow=visible)
        else:
            self.pres = self.ppt.Presentations.Open(
                                                    FileName = os.path.abspath(self.template), 
                                                    ReadOnly = True, 
                                                    Untitled = True, 
                                                    WithWindow = visible
                                                   )


    def add_slide(self, slide_num = None, layout = ppLayoutBlank, source_slide = None):
        """
        :param slide_num: Location to add new slide.  If not provided, new slide will be added at end of presentation
        :type slide_num: int
        :param layout: Base slide layout...need to understand this better...look up values online
        :type layout: int
        :param source_slide: Slide number to copy as a base for the new slide
        :type source_slide: int
        """        
        if slide_num is None:
            slide_num = self.pres.Slides.Count + 1
        
        if source_slide is None:
            self.pres.Slides.Add(Index = slide_num, Layout = layout)
        else:
            print source_slide
            self.pres.Slides(source_slide).Copy()
            self.pres.Slides.Paste(Index = slide_num)

    def add_picture(self, slide_num, fname, pic_format):
        """
        :param slide_num: Slide number to insert pictures
        :type slide_num: int
        :param fname: Filename of image to insert
        :type fname: string
        :param pic_format: Picture format object to define location, size, and crop
        :type pic_format: PictureFormat
        """

        pic = self.pres.Slides(slide_num).Shapes.AddPicture(  
                                                    FileName = os.path.abspath(fname), 
                                                    LinkToFile = False,
                                                    SaveWithDocument = True,
                                                    Left = pic_format.left, 
                                                    Top = pic_format.top,
                                                    #Width = pic_format.width,
                                                    #Height = pic_format.height
                                                 )          
        
        if pic_format.width is not None and pic_format.height is not None:
            pic.LockAspectRatio = False
        
        if pic_format.width is not None:
            pic.Width = pic_format.width
        
        if pic_format.height is not None:
            pic.Height = pic_format.height
        
        pic.PictureFormat.CropLeft = pic_format.crop_left
        pic.PictureFormat.CropTop = pic_format.crop_top
        pic.PictureFormat.CropRight = pic_format.crop_right
        pic.PictureFormat.CropBottom = pic_format.crop_bottom

    def add_textbox(self, slide_num, text, text_format):
        """
        :param slide_num: Slide number to insert pictures
        :type slide_num: int
        :param text: Text to show in Textbox
        :type text: string
        :param pic_format: Textbox format object to define location, size, and font style
        :type pic_format: TextboxFormat
        """
        textbox = self.pres.Slides(slide_num).Shapes.AddTextbox( 
                                                 Orientation = 1, # msoTextOrientationHorizontal
                                                 Left = text_format.left, 
                                                 Top = text_format.top,
                                                 Width = text_format.width,
                                                 Height = text_format.height
                                                 )         
        textbox.TextFrame.TextRange.Text = text
        textbox.TextFrame.TextRange.Font.Name = text_format.font_name
        textbox.TextFrame.TextRange.Font.Size = text_format.font_size
        textbox.TextFrame.TextRange.ParagraphFormat.Alignment = text_format.alignment
        textbox.TextFrame.TextRange.Font.Color = text_format.color
        textbox.TextFrame.TextRange.Font.Bold = text_format.bold
        textbox.TextFrame.TextRange.Font.Italic = text_format.italic
        textbox.TextFrame.TextRange.Font.Underline = text_format.underline


    def delete_slide(self, slide_num):
        """
        :param slide_num: Slide number to delete from presentation
        :type slide_num: int
        """
        self.pres.Slides(slide_num).Delete()
            
    def save(self):
        """
        Saves the current presentation.
        """
        self.pres.SaveAs(self.fname)
        
    def close(self):
        """
        Closes the current presentation.  Does not exit PowerPoint
        """
        self.pres.Close()
    
    def dump_pngs(self, outdir):
        """
        :param outdir: Folder name to export .png images of all slides into.  Will be created if does not exit
        :type outdir: string
        """        
        outdir = os.path.abspath(outdir)
        if not os.path.isdir(outdir):
            os.makedirs(outdir)
        self.pres.Export(Path = outdir, FilterName = 'PNG')


if __name__ == '__main__':
    #Development testing
    #my_pres = PPTPres(fname = 'Test.ppt', template = 'test_template.pptx', visible = True)
    my_pres = PPTPres(fname = 'Test.ppt', visible = False)    
    
    pic_format_1 = PictureFormat(left = 2, top = 3.3, width = 2, height = 3)    
    
    text_format_1 = TextboxFormat(left = 2, top = 1, width = 6, height = 2, alignment = alignCenter, font_size = 30)
    
    text_format_2 = TextboxFormat(left = 8, top = 6, width = 2, height = 1, alignment = alignCenter, font_size = 12, color = RGB(255,0,0))
    
    
    my_pres.add_slide()  
    my_pres.add_picture(slide_num = 1, fname = 'test_output/test_pic_1.jpg', pic_format = pic_format_1)  
    
    my_pres.add_slide()       
    #my_pres.add_slide()  
    
    #my_pres.add_slide(source_slide = 1)    
    
    #my_pres.delete_slide(slide_num = 2)  
    
    my_pres.add_textbox(slide_num = 1, text = 'I am a textbox', text_format = text_format_1)
    my_pres.add_textbox(slide_num = 1, text = 'Label', text_format = text_format_2)
    
    my_pres.save()
    my_pres.dump_pngs('test_output')
    #my_pres.close()
