import logging
log = logging.getLogger(__name__)

import macrohelper
basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa

from pythonize import wrapUnoContainer, UnoDateConverter

doc = basic.ThisComponent
text = doc.Text

from sys import platform

OPTIONAL_HYPHEN = chr(31)

def hyphdoc():
    if not platform.startswith('win'):
        basic.MsgBox("This macro requires Windows", "Bad platform")
        return
    import ctypes
    try:
        orfohym = ctypes.windll.orfo_hym
    except WindowsError:
        basic.MsgBox("This macro requires orfo_hym.dll in PATH",
                        "orfo_hym.dll missing") 
        return

    def hyphenate(word):
        hyphenated = ctypes.create_string_buffer(len(word))
        orfohym.HyphWord(word.encode('cp1251'), 
                hyphenated, 
                OPTIONAL_HYPHEN,
                0)
        return hyphenated

    cursor = text.createTextCursor()
    cursor.gotoStart(False)
    while True:
        cursor.gotoEndOfWord(True)
        hyphenated = hyphenate(cursor.String)
        cursor.setString(hyphenated)
        if not cursor.gotoNextWord(False):
            break
        
        
