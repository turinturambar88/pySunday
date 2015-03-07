from pySunday import slideBuilder




slides = slideBuilder.SundaySlides('01-25-15PM.pptx')

#How Great is our God
slides.add_scripture('Psalm 40')
slides.add_scripture('Isaiah 40:21-23')
slides.add_scripture('Malachi 3:1-5')
#Nothing But The Blood (307)
slides.add_hymn(307)
slides.add_scripture('Colossians 3:1-17')
#Give Thanks
slides.save()