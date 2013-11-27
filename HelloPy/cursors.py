"""Working with cursors:
    There are 2 type of cursors
    viewCursor = the real visible cursor, that knows all about current
    presentation of the text: pages, lines, screens ...
    textCursor = unvisible cursor that handles inner structure of the
    text, i.e. paragraphs, words, textparts...
"""

import logging
log = logging.getLogger(__name__)

import macrohelper
basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa

from pythonize import wrapUnoContainer, UnoDateConverter

doc = basic.ThisComponent
text = doc.Text


def InsertAtCursor(charcode = 0xA9):
    """insert character at the current cursor position"""
    vcursor = doc.CurrentController.ViewCursor
    text.insertString(vcursor.Start, chr(charcode), False)

def countStatistics():
    """ count text statistics using text cursor"""
    cursor = text.createTextCursor

    # count paragraphs
    cursor.gotoStart(False)
    paraCount = 0
    while cursor.gotoNextParagraph(False):
        paraCount += 1
    paraCount += 1

    # count sentences
    cursor.gotoStart(False)



