"""
Macroses and routines specific for Practica P.H. workflow
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
log = getLogger("pyuno.practica")

import writer
import hyphenate
from macrohelper import chars
# from unicodedata import name as uniname does not work on mac

# entity annotation template
_EAT = lambda char: \
    r"{Верстальщику: вставить символ, код символа = \u%04X}" % ord(char)
#    r"{Верстальщику: вставить символ %s, код символа = \u%04X}" % (
#        uniname(char), ord(char))


class VenturaPrepare:

    """
    several routines to prepare odt for Ventura markup
    """

    PATTERNS = (
        (r">", r">>"),  # angle brackets (less/greater then)
        (r"<", r"<<"),  # angle brackets (less/greater then)
        (r"(\d)\s(\d)", r"$1<|>$2"),  # DIGITAL_SPACE
        (chars.nonbreaking_space_code, r"<N>"),  # nonbreaking space
        # <space><emdash><space>
        (r"\s%s\s" % chars.mdash_code, "<N>%s " % chars.mdash),
        (r"\s%s\s" % chars.ndash_code, "<N>%s " % chars.ndash),
        (chars.soft_hyphen_code, r"<->"),  # SOFT_HYPHEN

    )

    SYMBOL_SUBST = (
        # tuple: (find range, basic deduction, exceptional mappings)
        (r"[\uF020-\uF0FF]", 0xF000, {}),  # ms symbol in priv use area (PUA)
        (r"[\u0391-\u03C9]", 0x350, {
            0x393: 0x47,
            0x396: 0x5a,  # etc. not completed
        })
    )
    SymbolFontName = "Symbol"
    VenturaEncoding = "cp1251"

    def __init__(self, basic, doc=None):
        self.doc = doc or basic.ThisComponent
        self.h = hyphenate.Hyphenate(self.doc,
                                     basic.GetDefaultContext())

        self.h.HyphMinWordLength = 4
        self.fru = writer.FindReplaceUtilities(self.doc)

    @staticmethod
    def make_subst(subst_tuple, char):
        _, deduction, mapping = subst_tuple
        char_code = ord(char)
        deducted = char_code - deduction
        log.debug("deducted = %s", deducted)
        ventura_entity = lambda x: "<@%03d>" % x
        return ventura_entity(mapping.get(char_code, deducted))

    def symbol_substitute(self):
        """
        substitute known chars with symbol analog
        """
        self.fru.SearchRegularExpression = True
        for ss in self.SYMBOL_SUBST:
            search = ss[0]
            for found in self.fru(search):
                char_entity = self.make_subst(ss, found.String)
                log.debug("char entity = %s", char_entity)
                found.String = char_entity
                found.CharFontName = self.SymbolFontName

    def unicode_annotate(self):
        """
        substitute unknown chars left of symbol_substitute
        insert editor comments and unichar name
        """
        from freq import construct_whitelist_search_range
        self.fru.SearchRegularExpression = True
        search = construct_whitelist_search_range(self.VenturaEncoding)
        for found in self.fru(search):
            char_entity = _EAT(found.String)
            found.String = char_entity

    def convert_index_markers(self):
        """
        convert bad im entries (index number like <111>) to
        good (index number like {111}
        """
        iu = writer.IndexUtilities2(self.doc)
        iu.convert_old_index_markers()

    def prepare_for_ventura(self):
        # self.convert_index_markers()  # temporarily
        self.h.hyphenate()

        # series of find-relacing routines
        # change hyphens
        self.fru.SearchRegularExpression = True
        for p in self.PATTERNS:
            self.fru(p[0], p[1])
        # substitute any non ascii character with corresponding Symbol char
        self.symbol_substitute()
        self.unicode_annotate()

    def __call__(self):
        self.prepare_for_ventura()


class BibliographyReorder:

    """
    Looks for bibliography markers aka [111]
    Finds appropriate records in bibliography list
    Inserts fields
    """
    BookmarkName = "bibliography"
    RecordPattern = r"(\d*)\.\s(.+)$"
    MarkerSearchPattern = r"\[(\d+,?\s?)+\]"

    def __init__(self, doc):
        self.doc = doc
        self.bu = writer.BookmarkUtilities(self.doc)
        self.cu = writer.CursorUtilities(self.doc)

    def get_bibliography_range(self):
        bookmarks = self.bu.getBookmarksDict()
        return bookmarks[self.BookmarkName].getAnchor()

    def make_biblist(self):
        import re
        pattern = re.compile(self.RecordPattern)
        bibliography = self.get_bibliography_range()
        biblist = []
        for p in self.cu.iterateParagraphs(bibliography):
            m = pattern.search(p.String)
            if m is not None:
                try:
                    num, rec = m.group(1, 2)
                except IndexError:
                    continue
                biblist.append((rec, num))
        return biblist

    def do_reorder(self):

        newbiblist = enumerate(sorted(self.make_biblist()), 1)

        # delete old bibliography
        bibcursor = self.get_bibliography_range()
        bibcursor.String = ""

        # change color for every [123] marker in text
        from macrohelper import colors
        fu = writer.FindReplaceUtilities(self.doc)
        fu.SearchRegularExpression = True
        fu.setReplaceAttributes(dict(CharColor=colors.red))
        fu(self.MarkerSearchPattern, "$0")

        # when replacing numbers change color back to default
        fu.setSearchAttributes(dict(CharColor=colors.red))
        fu.setReplaceAttributes(dict(CharColor=-1))

        tu = writer.TextUtilities(self.doc)
        for num, rectuple in newbiblist:
            rectitle, recnum = rectuple
            # replace markers
            fu(r"([\[\s])%s([,\]])" % recnum, "$1%s$2" % num)
            # print bibl record
            bibcursor = tu.insertParagraph(bibcursor, "%s.\t%s" % (
                num, rectitle))
            bibcursor.collapseToEnd()


def expand_table(basic, divider):
    """
    Converts pattern like @B ... $A plus table to list of rows compiled like
    for every word (splitted) by divider in column B in every row print
    B-word ... A-column
    @ - means expand
    $ - means substitude
    Was used for making drug list from Sanford
    """

    import re

    def print_string(pos=0, buf=''):
        if pos == len(parts):
            tu.appendPara(buf)
        else:
            if '@' in parts[pos] or '$' in parts[pos]:
                cell_contents = tbl.getCellByName(
                    "%s%s" % (parts[pos][-1], rownum)).String

                if '@' in parts[pos]:
                    _buf = buf
                    for sbst in cell_contents.split(divider):
                        buf = "%s %s" % (_buf, sbst)
                        print_string(pos + 1, buf)

                else:  # '$' in parts[pos]:
                    buf += cell_contents
                    print_string(pos+1, buf)
            else:
                buf += parts[pos]
                print_string(pos+1, buf)

    cu = writer.CursorUtilities(basic.ThisComponent)

    cur = cu.createTextCursorByANYRage(cu.getCurrentPosition())
    if cu.isInsideCell():
        output = basic.macro_create_doc("writer")
        tu = writer.TextUtilities(output)
        tbl = cur.TextTable
        pattern = tbl.getCellByPosition(0, 0).String.strip()
        parts = re.split("([$@][A-Z])", pattern)
        for rownum in range(2, tbl.Rows.Count + 1):
            print_string()
