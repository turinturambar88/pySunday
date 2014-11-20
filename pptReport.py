# -*- coding: utf-8 -*-
"""
Created on Wed Nov 19 21:16:49 2014

@author: Zach
"""


#Standard Library Imports
import os

#Anaconda Imports
import win32com.client


#Local Imports

inches_to_points = 72.0

#Text Alignment
alignLeft = 1 
alignCenter = 2
alignRight = 3


class PPTPres:
    """
    :param fname: Filename of output PowerPoint presentation
    :type fname: string
    """
    def __init__(self, fname, template = None, visible = True):
        self.fname = fname
        self.template = template
        
        #PowerPoint application
        self.ppt = win32com.client.Dispatch("PowerPoint.Application")
        
        if self.template is None:
            self.pres = self.ppt.Presentations.Add(WithWindow=visible)
        else:
            self.pres = self.ppt.Presentations.Open(
                                                    FileName = self.template, 
                                                    ReadOnly = True, 
                                                    Untitled = True, 
                                                    WithWindow = visible
                                                   )

    ppLayoutBlank = 12
    def add_slide(self, index = None, layout = ppLayoutBlank, source_slide = None):
        """
        """        
        if index is None:
            index = self.pres.Slides.Count + 1
        
        if source_slide is None:
            self.pres.Slides.Add(Index = index, Layout = layout)
        else:
            print source_slide
            self.pres.Slides(source_slide).Copy()
            self.pres.Slides.Paste(Index = index)

    def add_picture(self, index, fname, left, top, width, height, crop_left = 0, crop_top = 0, crop_right = 0, crop_bottom = 0):
        """
        :param index: Slide number to insert pictures
        :type index: int
        
        #Cropping is relative to image original size
        """

        pic = self.pres.Slides(index).Shapes.AddPicture(  
                                                    FileName = fname, 
                                                    LinkToFile = False,
                                                    SaveWithDocument = True,
                                                    Left = left * inches_to_points, 
                                                    Top = top * inches_to_points,
                                                    Width = width * inches_to_points,
                                                    Height = height * inches_to_points
                                                 )         
        pic.PictureFormat.CropLeft = crop_left * inches_to_points
        pic.PictureFormat.CropTop = crop_top * inches_to_points
        pic.PictureFormat.CropRight = crop_right * inches_to_points
        pic.PictureFormat.CropBottom = crop_bottom * inches_to_points

    def add_textbox(self, index, text, left, top, width, height, font_name = 'Arial', font_size = 12, alignment = alignLeft):
        """

        """
        textbox = self.pres.Slides(index).Shapes.AddTextbox( 
                                                 Orientation = 1, # msoTextOrientationHorizontal
                                                 Left = left * inches_to_points, 
                                                 Top = top * inches_to_points,
                                                 Width = width * inches_to_points,
                                                 Height = height * inches_to_points
                                                 )         
        textbox.TextFrame.TextRange.Text = text
        textbox.TextFrame.TextRange.Font.Name = font_name
        textbox.TextFrame.TextRange.Font.Size = font_size
        my_pres.pres.Slides(2).Shapes(3).TextFrame.TextRange.ParagraphFormat.Alignment = alignment


    def delete_slide(self, index):
        self.pres.Slides(index).Delete()
            
    def save(self):
        """
        """
        self.pres.SaveAs(self.fname)
        
    def close(self):
        """
        """
        self.pres.Close()
    
    def dump_pngs(self, outdir):
        """
        """        
        if not os.path.isdir(outdir):
            os.makedirs(outdir)
        self.pres.Export(Path = outdir, FilterName = 'PNG')


if __name__ == '__main__':
    #Development testing
    my_pres = PPTPres(fname = 'Test.ppt', template = 'test_template.pptx', visible = True)
    #my_pres = PPTPres(fname = 'Test.ppt', visible = True)    
    
    my_pres.add_slide()  
    
    my_pres.add_slide(source_slide = 1)    
    
    my_pres.delete_slide(index = 2)  
    
    my_pres.add_picture(index = 3, fname = 'test_output/Slide1.PNG', left = 2, top = 3.3, width = 1, height = 3, crop_right = 5)
    
    my_pres.add_textbox(index = 2, text = 'I am a textbox', left = 2, top = 3.3, width = 1, height = 3, alignment = alignRight)
    
    
    my_pres.save()
    my_pres.dump_pngs('test_output')
    #my_pres.close()