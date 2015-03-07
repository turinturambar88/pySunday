from pySunday import slideBuilder




slides = slideBuilder.SundaySlides('12-24-14PM.pptx')

slides.add_hymn(195)
slides.add_scripture('Jeremiah 33:14-18')
slides.add_scripture('Luke 2:1-7')
slides.add_hymn(230)
slides.add_scripture('Isaiah 41:1-4')
slides.add_scripture('Hebrews 4:12-13')
slides.add_hymn(211)
slides.add_scripture('Psalm 45:6-7')
slides.add_scripture('John 10:1-4,14-16')
slides.add_hymn(196)
slides.add_scripture('Psalm 132:13-18')
slides.add_scripture('John 19:1-5')
slides.add_hymn(296)
slides.add_scripture('Psalm 93:1-2')
slides.add_scripture('Revelation 19:11-16')
slides.add_hymn(193)
slides.add_scripture('Isaiah 9:6-7')
slides.add_scripture('1 Corinthians 15:21-26,56-57')
slides.add_hymn(203)


slides.save()