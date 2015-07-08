"""
Basic hyphenate routine using standard OO hyphenator, but placing
OPTIONAL_HYPHEN on every possible position, like orfohym does.
Usefull for exporting to Loyout

Copyright © 2015 Artem Putilov

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the "Software"),
to deal in the Software without restriction, including without limitation
the rights to use, copy, modify, merge, publish, distribute, sublicense,
and/or sell copies of the Software, and to permit persons to whom the
Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES
OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE
OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
"""
from logging import getLogger
log = getLogger("pyuno.hyphenate")
from com.sun.star.lang import Locale

from com.sun.star.text.ControlCharacter import SOFT_HYPHEN
from pythonize import wrapUnoContainer
import writer


class Hyphenate:
    """
    Does not implement Hyphenator service in OO. It just uses standard
    hyphenator to do its job
    """
    LmgrNS = "com.sun.star.linguistic2.LinguServiceManager"
    LpropNS = "com.sun.star.linguistic2.LinguProperties"
    LocaleTuple = ("ru", "RU", "")

    def __init__(self, doc, ctx):
        self.doc = doc
        self.ctx = ctx
        self.smgr = ctx.ServiceManager
        self.lmgr = self.smgr.createInstanceWithContext(self.LmgrNS, ctx)
        self.lprop = self.smgr.createInstanceWithContext(self.LpropNS, ctx)
        self.hyphenator = self.lmgr.getHyphenator()
        self.locale = Locale(*self.LocaleTuple)

    @property
    def HyphMinWordLength(self):
        return self.lprop.getPropertyValue("HyphMinWordLength")

    @HyphMinWordLength.setter
    def HyphMinWordLength(self, val):
        self.lprop.HyphMinWordLength = val

    def get_hyphenation_positions(self, word):
        log.debug("Hyphenating word %s", word)
        ph = self.hyphenator.createPossibleHyphens(word, self.locale, ())
        if ph is not None:
            log.debug("hyphenated word should be %s", ph.getPossibleHyphens())
            return ph.getHyphenationPositions()

    def hyphenate_text(self, cursor):
        cursor.gotoStart(False)
        while(True):
            while not(cursor.isStartOfWord()) and cursor.goRight(1, False):
                pass
            cursor.gotoEndOfWord(True)
            hpositions = self.get_hyphenation_positions(cursor.String)
            log.debug("hpositions is %s", hpositions)
            if hpositions is not None:
                for pos in reversed(hpositions):
                    cursor.gotoStartOfWord(False)
                    cursor.goRight(pos + 1, False)
                    self.insert_hyphen(cursor)
            if not cursor.gotoNextWord(False):
                break

    def insert_hyphen(self, cursor):
        cursor.Text.insertControlCharacter(
            cursor.Start, SOFT_HYPHEN, False)

    def hyphenate(self):
        # iterate words
        # get positions
        # move cursor on that positions in reverse order
        # insert control character SOFT_HYPHEN
        cursor = self.doc.Text.createTextCursor()
        self.hyphenate_text(cursor)

        # hyphenate all tables
        cu = writer.CursorUtilities(self.doc)
        for t in wrapUnoContainer(self.doc.getTextTables(), "XIndex"):
            for cell in cu.iterateTableCells(tbl=t):
                cursor = cell.Text.createTextCursor()
                self.hyphenate_text(cursor)


from sys import platform

OPTIONAL_HYPHEN = b'\x1f'  # b'\xad'  # chr(173)  # '\u00AD'.encode('cp1251')
EN_DASH = '-'


def hyphdoc():
    """
    Alternative hyphenator for windows only. Uses orpho_hyph.dll
    """

    basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa
    doc = basic.ThisComponent
    text = doc.Text
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
            log.debug("Hyphenated == Original: %s",
                      hyphenated == cursor.String)
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
