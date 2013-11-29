""" working with character properties """

import logging
log = logging.getLogger(__name__)

import macrohelper
basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa

from pythonize import wrapUnoContainer, UnoDateConverter

doc = basic.ThisComponent
text = doc.Text

def changeFontRelief():
    paraCount = 0
    for para in wrapUnoContainer(text):
        paraCount += 1
        para.CharRelief = paraCount % 3

def removeFontRelief():
    from com.sun.star.text.FontRelief import NONE as NORELIEF
    paraCount = 0
    for para in wrapUnoContainer(text):
        paraCount += 1
        para.CharRelief = NORELIEF



