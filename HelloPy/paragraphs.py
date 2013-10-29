"""working with paragraphs"""

import logging
log = logging.getLogger(__name__)

from tema.oo import macrohelper
basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa

from tema.oo.pythonize import wrapUnoContainer, UnoDateConverter

doc = basic.ThisComponent
text = doc.Text


def enumerateParagraphs():
    line = "This doc contains {} paragraphs and {} tables"
    paraCount = tableCount = 0
    from com.sun.star.style.ParagraphAdjust import RIGHT  
    for para in wrapUnoContainer(text):  # text has XEnumerateAccess interface
        log.debug("para is %s", para)
        log.debug(dir(para))

        if para.supportsService("com.sun.star.text.Paragraph"):
            paraCount += 1 
            if not paraCount % 5:
                para.ParaAdjust = RIGHT

                # this is how we work with structures: we copy or instantiate
                # them and then assign to the property
                from com.sun.star.style import LineSpacing
                from com.sun.star.style.LineSpacingMode import LEADING
                # we could say linespacing = para.ParaLineSpacing
                linespacing = LineSpacing()  
                linespacing.Mode = LEADING # just int constant
                linespacing.Height = 1000
                para.ParaLineSpacing = linespacing
        if para.supportsService("com.sun.star.text.Table"):
            tableCount += 1 

    basic.MsgBox(line.format(paraCount, tableCount), "Statistics")


def enumerateTextSections():
    """Just as you can enumerate the paragraphs in a document,
    you can enumerate the text sections in a paragraph.
    Text within each enumerated portion uses the same
    properties and is of the same type."""

    paraCount = 0
    msg = ""
    for para in wrapUnoContainer(text):
        paraCount += 1
        msg += "{}. ".format(paraCount)
        for textpart in wrapUnoContainer(para):
            msg += "{}:".format(textpart.TextPortionType)
        msg += "\n"
        if not paraCount % 10:
            basic.MsgBox(msg, "Paragraph text sections")
            break


