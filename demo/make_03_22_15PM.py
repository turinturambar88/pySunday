from pySunday import slideBuilder

slides = slideBuilder.SundaySlides('2015-03-22PM.pptx')

slides.add_hymn(4)
slides.add_scripture('Psalm 5')
slides.add_scripture('Revelation 1:8')
slides.add_scripture('Romans 5')
#I Will Sing of My Redeemer
slides.add_scripture('Hebrews 3:7-19')
#Ancient of Days
slides.save()