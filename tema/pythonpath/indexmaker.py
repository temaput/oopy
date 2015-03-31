"""
Generates index given an output from layout software Ventura
Copyright (C) 2015 Artem Putilov

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
log = getLogger("pyuno.indexmaker")
from collections import namedtuple
from macrohelper import colors, chars
import writer


from re import compile
IndexLeaf = namedtuple("IndexLeaf", "pageSet, subLevelsDct")
MatchLine = namedtuple("MatchLine", "entry1, entry2, entry3, diapasonMarker,"
                       "counter, page")

indexSigns = writer.IndexUtilities.signs
MaxLevels = writer.IndexUtilities.MaxLevels

context = dict(
    entry="[^:=+{]+",
    diapasonMarker="(?P<diapasonMarker>[+=])",
    counter="{(?P<counter>\d+)}",
    page="(?P<pageNumber>\d+)"
)
template = (
    r"^(?P<entry1>{entry}):?"
    r"(?P<entry2>{entry})?:?"
    r"(?P<entry3>{entry})?{diapasonMarker}?{counter}\t{page}"
)

r = compile(template.format(**context))


def appendPageNum(numbersList, pn, prev_pn, range_start,
                  range_delimiter=chars.mdash):
        # we need to print something
        if range_start is None:
            # print single page number
            numbersList.append(str(prev_pn))
        else:
            # print page diapason
            if prev_pn - range_start == 1:
                #  short range like 123, 124
                numbersList.append(str(range_start))
                numbersList.append(str(prev_pn))
            else:
                #  long range like 123-129
                numbersList.append(
                    "{range_start}{delimiter}{range_stop}".format(
                        **dict(range_start=range_start,
                               range_stop=prev_pn,
                               delimiter=range_delimiter)
                    ))


def print_page_set(pageSet, range_delimiter=chars.mdash,
                   list_delimiter=", "):
    numbersList = []
    prev_pn = None
    range_start = None
    for pn in sorted(pageSet):
        if prev_pn is not None:
            if pn - prev_pn > 1:
                appendPageNum(numbersList, pn, prev_pn, range_start)
                range_start = None
                prev_pn = None
            else:
                # wait for the range to stop
                if range_start is None:
                    range_start = prev_pn
        prev_pn = pn
    # print last portion
    if prev_pn is not None:
        appendPageNum(numbersList, pn, prev_pn, range_start)
    return list_delimiter.join(numbersList)


class BadIndexEntries(ValueError):
    pass


class IndexMaker:

    MarginStep = 500
    StyleNames = ("i01", "i02", "i03")

    def __init__(self, in_doc, out_doc):
        self.doc = in_doc
        self.output_document = out_doc
        self.Text = writer.TextUtilities(self.doc)
        self.Cursor = writer.CursorUtilities(self.doc)
        self.Styles = writer.StyleUtilities(self.output_document)

    def createIndexStyles(self):
        for i in range(3):
            self.Styles.createParaStyle(
                self.StyleNames[i], {"ParaLeftMargin": self.MarginStep * i})

    def markUnmatchedEntries(self):
        i = 0
        for p in self.Cursor.iterateParagraphs():
            if len(p.String) and r.match(p.String) is None:
                i += 1
                self.Text.markRange(p, colors.yellow)
        return i

    def collectMatches(self):
        matches = {}
        for p in self.Cursor.iterateParagraphs():
            m = r.match(p.String)
            if m is not None:
                match_line = MatchLine(*m.groups())
                matches[match_line.counter] = match_line
        return matches

    def parseMatches(self, matches):
        indexTree = {}
        while(matches):
            item = matches.popitem()
            mline = item[1]
            # check for page range
            branch = self.getBranch(mline, indexTree)

            counter = int(mline.counter)
            page = int(mline.page)
            pageRange = (page,)
            # check for diapason
            if mline.diapasonMarker is not None:
                if mline.diapasonMarker == indexSigns.diapasonOpening:
                    lookup = counter + 1
                    lookupMarker = indexSigns.diapasonClosing
                    rangeList = [page, None]
                else:
                    lookup = counter - 1
                    lookupMarker = indexSigns.diapasonOpening
                    rangeList = [None, page + 1]
                lookup = str(lookup)
                if (lookup in matches and
                        matches[lookup].diapasonMarker == lookupMarker):
                    opposite_mline = matches.pop(lookup)
                    if rangeList[1] is None:
                        rangeList[1] = int(opposite_mline.page) + 1
                    else:
                        rangeList[0] = int(opposite_mline.page)
                    pageRange = tuple(range(*rangeList))
                else:
                    # diapason closing entry not found
                    log.warning("Closing marker not found with counter=%s",
                                lookup)
            # fill the branch
            branch.pageSet.update(pageRange)
        return indexTree

    def printTree(self, indexTree, level=0):
        for bname in sorted(indexTree):
            branch = indexTree[bname]
            print ("%s%s%s  %s" % (level, level * "\t", bname,
                                   print_page_set(branch.pageSet)))
            if len(branch.subLevelsDct):
                self.printTree(branch.subLevelsDct, level + 1)

    def printTreeToDoc(self, cur, indexTree, level=0):
        for bname in sorted(indexTree):
            branch = indexTree[bname]
            newPara = self.Text.insertParagraph(
                cur, "%s  %s" % (bname, print_page_set(branch.pageSet)))
            newPara.ParaStyleName = self.StyleNames[level]
            newPara.collapseToEnd()
            if len(branch.subLevelsDct):
                self.printTreeToDoc(newPara, branch.subLevelsDct, level + 1)

    def makeIndex(self):
        unmatchedEntries = self.markUnmatchedEntries()
        if unmatchedEntries > 0:
            raise BadIndexEntries(
                "There are %s unmatched entries" % unmatchedEntries
            )
        indexTree = self.parseMatches(self.collectMatches())
        self.Text = writer.TextUtilities(self.output_document)
        self.createIndexStyles()
        self.printTreeToDoc(self.output_document.Text.Start,
                            indexTree)

    def getBranch(self, mline, branchDct, level=0):
        entry = mline[level]
        if entry not in branchDct:
            # this is a new branch
            branchDct[entry] = IndexLeaf(set(), {})
        branch = branchDct[entry]
        if level < MaxLevels - 1 and mline[level+1] is not None:
            return self.getBranch(mline, branch.subLevelsDct, level + 1)
        else:
            return branch
