from pySunday import slideBuilder




slides = slideBuilder.SundaySlides('01-11-15PM.pptx')

slides.add_scripture('Psalm 19')
slides.add_scripture('John 4:24')
slides.add_scripture('2 Corinthians 5:1-10')
slides.add_scripture('1 Peter 3:8-17')

slides.save()