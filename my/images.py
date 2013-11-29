"""working with graphics"""


import logging
log = logging.getLogger(__name__)

import macrohelper
basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa

from pythonize import wrapUnoContainer, UnoDateConverter

doc = basic.ThisComponent
text = doc.Text

def insertImage():
    url = basic.InputBox("Please type image URL", "Path")
    if url:
        insertImageByURL(doc, url)

def insertImageByURL(doc, url):
    import unohelper
    text = doc.Text
    cursor = text.createTextCursor()
    cursor.goToStart(False)
    img = doc.createInstance("com.sun.star.text.GraphicObject")
    img.GraphicURL = url
    img.Width = 6000
    img.Height = 8000
    img.AnchorType = unohelper.uno.getConstantByName(
            "com.sun.star.text.TextContentAnchorType.AS_CHARACTER")

    text.insertTextContent(cursor, img, False)
    
