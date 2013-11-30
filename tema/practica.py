import logging
log = logging.getLogger(__name__)

import macrohelper
basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa

from pythonize import wrapUnoContainer, UnoDateConverter

doc = basic.ThisComponent
text = doc.Text

from sys import platform

OPTIONAL_HYPHEN = b'\x1f' #  b'\xad'  # chr(173)  # '\u00AD'.encode('cp1251')
EN_DASH = '-'

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


    def cycle_hyphenation(cursor):

        def hyphenate(word):
            en_pos = word.find(EN_DASH)
            log.debug('en_pos=%s', en_pos)
            if en_pos > -1:
                log.debug('before dash: %s', word[:en_pos])
                log.debug('afterdash: %s', word[en_pos+1:])
                return '{}{}{}'.format(
                        hyphenate(word[:en_pos]),
                        EN_DASH,
                        hyphenate(word[en_pos+1:])
                        )
            hyphenated = ctypes.create_string_buffer(len(word)*2)
            try:
                _word = word.encode('cp1251')
            except UnicodeEncodeError:
                return word
            orfohym.HyphWord(_word, 
                    hyphenated, 
                    OPTIONAL_HYPHEN,
                    0)
            return hyphenated.value.replace(b'\x1f', b'\xad').decode('cp1251')

        while True:
            cursor.gotoEndOfWord(True)
            if cursor.String.endswith('.'):
                cursor.goLeft(1, True)
            hyphenated = hyphenate(cursor.String)
            log.debug(20 * '- - ')
            log.debug("Original: %s", cursor.String)
            log.debug("Hyphenated: %s", hyphenated)
            log.debug("Hyphenated == Original: %s", hyphenated == cursor.String)
            if hyphenated != cursor.String:
                cursor.setString(hyphenated)
            if not cursor.gotoNextWord(False):
                break

    cursor = text.createTextCursor()
    tables = wrapUnoContainer(doc.getTextTables())
    cursor.gotoStart(False)
    cycle_hyphenation(cursor)
        
    for tkey in tables:
        t = tables[tkey]
        for cell_name in t.getCellNames():
            cell = t.getCellByName(cell_name)
            cursor = cell.Text.getStart()
            cycle_hyphenation(cursor)
