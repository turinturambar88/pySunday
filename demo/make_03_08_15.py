from pySunday import slideBuilder




slides = slideBuilder.SundaySlides('2015-03-08PM.pptx')

slides.add_hymn(38)
slides.add_scripture('Psalm 19')
slides.add_scripture('James 3:13-17')
slides.add_scripture('John 6:35-58')
#You are my King (Amazing Love)
slides.add_scripture('Psalm 90')
#Hallelujah, What a Savior 
slides.save()