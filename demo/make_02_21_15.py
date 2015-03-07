from pySunday import slideBuilder




slides = slideBuilder.SundaySlides('02-22-15PM.pptx')

#God You Reign
slides.add_scripture('Hebrews 1:1-14')
slides.add_scripture('Malachi 3:6')
slides.add_scripture('Isaiah 1:11-20')
slides.add_hymn(402)
slides.add_scripture('Joshua 24:14-28')
#O Church Arise
slides.save()