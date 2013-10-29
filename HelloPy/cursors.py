"""Working with cursors:
    There are 2 type of cursors
    viewCursor = the real visible cursor, that knows all about current
    presentation of the text: pages, lines, screens ...
    textCursor = unvisible cursor that handles inner structure of the
    text, i.e. paragraphs, words, textparts...
"""

import logging
log = logging.getLogger(__name__)

from tema.oo import macrohelper
basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa

from tema.oo.pythonize import wrapUnoContainer, UnoDateConverter

doc = basic.ThisComponent
text = doc.Text


def InsertAtCursor(charcode = 0xA9):
    vcursor = doc.CurrentController.ViewCursor
    text.insertString(vcursor.Start, chr(charcode), False)

