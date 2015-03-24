"""
Generic classes for working with LibreOffice Writer objects
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

import logging
log = logging.getLogger('pyuno.writer')

from pythonize import wrapUnoContainer

from com.sun.star.text.ControlCharacter import (  # noqa
                                                PARAGRAPH_BREAK,
                                                LINE_BREAK,
                                                HARD_HYPHEN,
                                                SOFT_HYPHEN,
                                                HARD_SPACE,
                                                APPEND_PARAGRAPH)


class BadSelection(ValueError):
    pass


class BaseUtilities:
    def __init__(self, _doc):
        self.doc = _doc


class Properties:
    """
    Work with properties
    """

    @classmethod
    def setFromDict(cls, obj, dct):
        for k in dct:
            obj.setPropertyValue(k, dct[k])

    @classmethod
    def dctFromProperties(cls, obj):
        PropertySetInfo = obj.getPropertySetInfo()
        dct = {}
        for p in PropertySetInfo.Properties:
            dct[p.Name] = obj.getPropertyValue(p.Name)
        return dct


class TextUtilities(BaseUtilities):
    """
    Work with text, paragraphs
    """

    def appendText(self, t):
        self.doc.Text.insertString(
            self.doc.Text.getEnd(),
            t,
            False)

    def appendPara(self, t):
        self.appendText(t)
        self.doc.Text.insertControlCharacter(
            self.doc.Text.getEnd(), PARAGRAPH_BREAK, False)


class IndexUtilities(BaseUtilities):
    """
    Short routines for manipulating index marks in text
    """
    IndexNS = "com.sun.star.text.DocumentIndex"
    MarksContainerProperty = "DocumentIndexMarks"
    MarkNS = "com.sun.star.text.DocumentIndexMark"
    MarkKeySeparator = ":"
    MarkKeyNames = ('AlternativeText', 'PrimaryKey', 'SecondaryKey')
    MarkPresentationMask = "{XE %s}"
    ShowMarksVarName = "showIndexMarks"
    MarkNumberBrackets = ("<", ">")
    MarkPropertiesTemplate = {"AlternativeText": "",
                              "PrimaryKey": "",
                              "SecondaryKey": ""}
    RefCount = 0

    def __del__(self):
        log.debug("Remove instance of IndexUtilities")

    def __init__(self, *args, **kwargs):
        super(IndexUtilities, self).__init__(*args, **kwargs)
        self.Cursor = CursorUtilities(self.doc)
        self.Fields = FieldUtilities(self.doc)
        self.FirstEntryList = self.AlternativeTextList = []
        self.SecondEntryList = self.PrimaryKeyList = []
        self.ThirdEntryList = self.SecondaryKeyList = []
        self.LastMarkNum = None
        self.MarkCacheDict = {}
        self.incrementRefCount()
        log.debug("initializing %s instance of IndexUtitlities", self.RefCount)

    @classmethod
    def incrementRefCount(cls):
        cls.RefCount += 1

    def iterMarks(self, rng):
        """
        iterate over indexMarks in range
        """
        for portion in self.Cursor.iterateTextPortions(rng):
            if portion.TextPortionType == "DocumentIndexMark":
                yield portion.DocumentIndexMark

    def keysFromString(self, markString):
        markProperties = self.MarkPropertiesTemplate.copy()
        markProperties.update({k: v for (k, v) in zip(
            self.MarkKeyNames,
            markString.split(self.MarkKeySeparator)
        )})
        return markProperties

    def keysToList(self, markProperties):
        markKeysList = []
        for k in self.MarkKeyNames:
            if k in markProperties:
                keyField = markProperties[k]
                if len(keyField):
                    markKeysList.append(keyField)
        return markKeysList

    def keysToString(self, markProperties):
        """
        reverse of indexMarkKeysFromString
        """
        return self.MarkKeySeparator.join(self.keysToList(markProperties))

    def makeMarkPresentation(self, mark):
        """
        Prepares mark presentation for hidden text field
        aka {"XE" Name:First key:SecondaryKey}
        """
        markProperties = Properties.dctFromProperties(mark)
        return self.MarkPresentationMask % self.keysToString(
            markProperties)

    def makeMarkHere(self, markString):
        """
        Check for selection if it lies acros more than 1 paragraph mark range
        """
        if self.LastMarkNum is None:
            self.rebuildCache()

        lastMarkNum = self.LastMarkNum
        if self.Cursor.isInsideParagraph() or self.Cursor.isInsideCell():
            markString = "%s<%s>" % (markString, lastMarkNum)
            markProperties = self.keysFromString(markString)
            mark = self.createMark(markProperties)
            self.insertMark(mark)
            log.debug("mark inserted, LMN=%s", self.LastMarkNum)
            self.addMarkToCache(mark, lastMarkNum)
            self.addMarkKeysToCache(markProperties)
            log.debug("cache updated")
            self.LastMarkNum += 1
            log.debug("cache updated, LMN=%s", self.LastMarkNum)
        else:
            # more than one paragraph selected
            selEdges = self.Cursor.getSelectionEdges()
            if selEdges is None:
                raise BadSelection("Please, select something else")
            # Places 2 index marks at the start and end of textRange
            markProperties = self.keysFromString(markString)
            self.addMarkKeysToCache(markProperties)
            mark1 = self.createMark(
                self.keysFromString("%s+<%s>" % (markString, lastMarkNum)))
            mark2 = self.createMark(
                self.keysFromString("%s=<%s>" % (markString, lastMarkNum + 1)))
            self.insertMark(mark1, selEdges[0])
            self.insertMark(mark2, selEdges[1], False)
            self.addMarkToCache(mark1, lastMarkNum)
            self.addMarkToCache(mark2, lastMarkNum + 1)
            self.LastMarkNum += 2
        self.doc.TextFields.refresh()

    def stripMarkNumber(self, keyString):
        """
        remove =<123> or <123> or +<123> from the end of keyString
        """
        sepPosition = len(keyString)
        for sep in ('+', '=', '<'):
            if sep in keyString:
                sepPosition = min(sepPosition, keyString.find(sep))
        return keyString[:sepPosition]

    def addMarkToCache(self, mark, markNumber):
        log.debug("add mark number %s to cache", markNumber)
        self.MarkCacheDict[markNumber] = mark

    def rebuildCache(self):
        marks = self.getMarks()
        self.LastMarkNum = len(marks)
        for m in marks:
            markProperties = Properties.dctFromProperties(m)
            self.addMarkKeysToCache(markProperties)
            markKeysList = self.keysToList(markProperties)
            markNumber = self.readMarkNumber(markKeysList)
            if markNumber is not None:
                self.LastMarkNum = max(self.LastMarkNum, markNumber)
                self.addMarkToCache(m, markNumber)

    def addMarkKeysToCache(self, markProperties):
        for keyName in self.MarkKeyNames:
            keyValue = self.stripMarkNumber(markProperties[keyName])
            if keyValue:
                keyNameList = getattr(self, "%sList" % keyName)
                if keyValue not in keyNameList:
                    keyNameList.append(keyValue)

    def removeMarkHere(self):
        """
        Check current selection for any presentations
        and get rid of them and their marks
        """
        if self.LastMarkNum is None:
            self.rebuildCache()

        marksRemoved = 0
        # if something is selected, look for all fields in selection
        if self.Cursor.isAnythingSelected():
            for selectionCursor in self.Cursor.iterateSelections():
                for textField in self.Fields.iterateTextFields(
                        selectionCursor):
                    # check that it is our field
                    log.debug("Checking that %s is in %s",
                              self.ShowMarksVarName,
                              textField.Condition)
                    if self.ShowMarksVarName in textField.Condition:
                        # find attached index mark
                        for mark in self.getAttachedIndexMarks(textField):
                            for m in self.getLinkedMarks(mark):
                                log.debug("Marks removed %s, now remove %s",
                                          marksRemoved, m)
                                marksRemoved += 1
                                m.dispose()
                        # remove presentation aswell
                        textField.dispose()
        return marksRemoved

    def getAttachedIndexMarks(self, presentationField):
        cur = presentationField.Anchor.Text.createTextCursorByRange(
            presentationField.Anchor)
        cur.goRight(2, True)
        log.debug("Found index marks under cursor %s are %s", cur.String,
                  [im for im in self.iterMarks(cur)])
        return self.iterMarks(cur)

    def getAttachedPresentationFields(self, mark):
        """
        travel cursor backwards to find the closest presentation field
        """
        pass

    def getLinkedMarks(self, mark):
        """
        Inspects if it is a diapason marker
        returns tuple with mark and its opposite (if exist)
        """
        markKeysList = self.keysToList(Properties.dctFromProperties(mark))
        markString = markKeysList[-1]  # we only check the last element
        if '=' in markString or '+' in markString:
            markNumber = self.readMarkNumber(markKeysList)
            if markNumber is not None:
                mark2 = self.getMarkByNumber(markNumber + 1)
                if mark2 is not None:
                    return (mark, mark2)
        # this is single mark
        return (mark,)

    def getMarkByNumber(self, markNumber):
        log.debug("is mark number %s in cache? %s",
                  markNumber, markNumber in self.MarkCacheDict)
        if markNumber in self.MarkCacheDict:
            return self.MarkCacheDict[markNumber]

    def readMarkNumber(self, markKeysList):
        markString = markKeysList[-1]  # we only look in last entry
        bracket1 = markString.find(self.MarkNumberBrackets[0]) + 1
        bracket2 = markString.find(self.MarkNumberBrackets[1])
        try:
            return(int(markString[bracket1:bracket2]))
        except ValueError:
            pass

    def toggleMarks(self, toggle=None):
        self.Fields.toggleHiddenText(self.ShowMarksVarName, toggle)

    def createMark(self, markProperties):
        """
        Creates mark and fills its properties:
        - AlternativeText
        - PrimaryKey
        - SecondaryKey
        - IsMainEntry
        """

        mark = self.doc.createInstance(self.MarkNS)
        Properties.setFromDict(mark, markProperties)
        return mark

    def insertMark(self, mark, insertionPoint=None, givePresentation=True):
        """
        inserts index mark at the current cursor position or specified range
        prepends it with hiddentext field containing mark presentation
        """
        if insertionPoint is None:
            insertionPoint = self.Cursor.getCurrentPosition()
        cur = insertionPoint.Text.createTextCursorByRange(insertionPoint)
        cur.Text.insertTextContent(cur,
                                   mark, False)
        if givePresentation:
            field = self.Fields.createHiddenTextField(
                self.makeMarkPresentation(mark), varName=self.ShowMarksVarName)

            cur.goLeft(1, False)
            cur.Text.insertTextContent(cur,
                                       field, False)

    def getMarks(self):
        """
        Creates tuple of document index marks
        """
        DocumentIndex = self.getDocumentIndex()
        return DocumentIndex.getPropertyValue(self.MarksContainerProperty)

    def getDocumentIndex(self):
        """
        gets whole document index
        """
        return self.doc.createInstance(self.IndexNS)


class IndexUtilities2(IndexUtilities):
    """
    Takes different approach to index mark keys
    AlternativeText always contains last entry

    AlternativeText
    PrimaryKey:AlternativeText
    PrimaryKey:SecondaryKey:AlternativeText
    """

    def keysFromString(self, markString):
        markTuple = [k.strip() for k in
                     markString.split(self.MarkKeySeparator)]
        markKeys = self.MarkPropertiesTemplate.copy()
        markKeys['AlternativeText'] = markTuple[-1]
        if len(markTuple) > 1:
            markKeys['PrimaryKey'] = markTuple[0]
            if len(markTuple) > 2:
                markKeys['SecondaryKey'] = markTuple[1]
        return markKeys

    def keysToList(self, markProperties):
        markKeysList = [markProperties['AlternativeText']]
        if len(markProperties['PrimaryKey']):
            markKeysList.insert(0, markProperties['PrimaryKey'])
            if len(markProperties['SecondaryKey']):
                markKeysList.insert(1, markProperties['SecondaryKey'])
        return markKeysList

    def addMarkKeysToCache(self, markProperties):
        log.debug("in iu2 addMarkKeysToCache")
        markKeysList = self.keysToList(markProperties)
        log.debug("adding mark keys to Cache: %s", markKeysList)
        lastEntry = markKeysList[-1]
        markKeysList[-1] = self.stripMarkNumber(lastEntry)

        if markKeysList[0] not in self.FirstEntryList:
            self.FirstEntryList.append(markKeysList[0])
        if len(markKeysList) > 1 \
                and markKeysList[1] not in self.SecondEntryList:
            self.SecondEntryList.append(markKeysList[1])
        if len(markKeysList) > 2 \
                and markKeysList[2] not in self.ThirdEntryList:
            self.ThirdEntryList.append(markKeysList[2])

        log.debug("FirstEntryList len = %s", len(self.FirstEntryList))
        log.debug("SecondEntryList len = %s", len(self.SecondEntryList))


class FieldUtilities(BaseUtilities):
    """
    Work with fields
    """
    branch = "com.sun.star.text.FieldMaster.User"
    FieldMasterNS = "com.sun.star.text.FieldMaster"
    TextFieldNS = "com.sun.star.text.TextField"

    def __init__(self, *args, **kwargs):
        super(FieldUtilities, self).__init__(*args, **kwargs)
        if 'FieldMasterType' in kwargs:
            self.branch = "%s.%s" % (
                self.FieldMasterNS, kwargs['FieldMasterType'])
        from com.sun.star.text.SetVariableType import (
            VAR,)
        self.VAR = VAR
        self.Cursor = CursorUtilities(self.doc)

    def iterateTextFields(self, rng):
        """
        Iterate text fields in rng
        """
        for portion in self.Cursor.iterateTextPortions(rng):
            if portion.TextPortionType == "TextField":
                yield portion.TextField

    def getOrCreateMaster(self, Name, FieldMasterType=None,
                          FieldMasterProperties={}):
        """
        looks for MasterField by name
        if found returns it, or creates it if not
        """
        if FieldMasterType is not None:
            branch = "%s.%s" % (
                self.FieldMasterNS, FieldMasterType)
        else:
            branch = self.branch

        fullname = "%s.%s" % (branch, Name)
        fmasters = self.doc.getTextFieldMasters()

        if fmasters.hasByName(fullname):
            return fmasters.getByName(fullname)
        else:
            fmaster = self.doc.createInstance(branch)
            fmaster.Name = Name
            Properties.setFromDict(fmaster, FieldMasterProperties)
            return fmaster

    def createTextField(self, textFieldType, fieldProperties={}):
        branch = "%s.%s" % (self.TextFieldNS, textFieldType)
        textField = self.doc.createInstance(branch)
        if fieldProperties:
            Properties.setFromDict(textField, fieldProperties)
        return textField

    def toggleHiddenText(self, varName="showHidden", varContent=None):
        """
        Create or check for variable at the begining of doc
        it will define visibility of hidden field
        """
        fmaster = self.getOrCreateMaster(varName, "SetExpression",
                                         {"SubType": self.VAR})
        if len(fmaster.DependentTextFields):
            dependant = fmaster.DependentTextFields[0]
            if varContent is None:
                varContent = 1 if dependant.Content == '0' else 0
            self.doc.Text.removeTextContent(dependant)

        dependant = self.createTextField("SetExpression")
        dependant.IsVisible = False
        dependant.attachTextFieldMaster(fmaster)
        dependant.Content = 1 if varContent is None else varContent
        self.doc.Text.insertTextContent(
            self.doc.Text.getStart(), dependant, False)
        self.doc.TextFields.refresh()

    def createHiddenTextField(self,
                              Content, Condition=None, varName="showHidden"):
        Condition = Condition or "%s != 1" % varName
        self.toggleHiddenText(varName, 1)
        return self.createTextField(
            "HiddenText", {"Condition": Condition, "Content": Content})


class CursorUtilities(BaseUtilities):
    """
    Work with view and text cursors, selection and ranges
    """

    TableCursorNS = "com.sun.star.text.TextTableCursor"
    TextRangesNS = "com.sun.star.text.TextRanges"
    TableRangeNameSplitter = ":"

    def getViewCursor(self):
        return self.doc.CurrentController.getViewCursor()

    def getCurrentPosition(self):
        viewCursor = self.getViewCursor()
        return viewCursor.getStart()

    def createTextCursorByEdges(self, startRange, endRange):
        cursor = self.doc.Text.createTextCursor()
        cursor.gotoRange(startRange, False)
        cursor.gotoRange(endRange, True)
        return cursor

    def getEdge(self, rng, edge="left"):
        cur = rng.Text.createTextCursorByRange(rng)
        {"left": cur.collapseToStart, "right": cur.collapseToEnd}[edge]()
        return cur

    def forwards(self, rng):
        return self.createTextCursorByEdges(
            self.getEdge(rng), self.getEdge(rng, "right"))

    def backwards(self, rng):
        return self.createTextCursorByEdges(
            self.getEdge(rng, "right"), self.getEdge(rng))

    def iterateTableCells(self, cellRangeName=None, tbl=None):
        """
        gets table obj and iterates within its cells
        """
        opened = False
        closed = False
        if tbl is None:
            selections = self.doc.getCurrentSelection()
            if selections is not None:
                if selections.supportsService(self.TableCursorNS):
                    vc = self.getViewCursor()
                    tbl = vc.TextTable
                    if cellRangeName is None:
                        cellRangeName = selections.RangeName
                elif selections.supportsService(self.TextRangesNS):
                    sel = selections.getByIndex(0)
                    tbl = sel.TextTable
        if tbl is not None and cellRangeName is not None:
            cellRange = cellRangeName.split(self.TableRangeNameSplitter)
            if len(cellRange) == 1:
                yield tbl.getCellByName(cellRange[0])
            else:
                for cellName in tbl.getCellNames():
                    if cellName == cellRange[0]:
                        opened = True
                    if cellName == cellRange[1]:
                        closed = True
                    if opened and not closed:
                        yield tbl.getCellByName(cellName)

    def iterateParagraphs(self, rng):
        return wrapUnoContainer(rng)

    def iterateTextPortions(self, rng):
        if self.isInsideCell(rng):
            enclosingRng = rng.Cell
        else:
            enclosingRng = rng
        for para in self.iterateParagraphs(enclosingRng):
            for portion in wrapUnoContainer(para):
                if self.isOverlaping(portion, rng):
                    yield portion

    def iterateSelections(self):
        """
        iterates over selected fragments and returns iterator of textCursors
        """
        selections = self.doc.getCurrentSelection()
        if selections is not None:
            if selections.supportsService(self.TableCursorNS):
                vc = self.getViewCursor()
                tbl = vc.TextTable
                cellRangeName = selections.RangeName
                for cell in self.iterateTableCells(cellRangeName, tbl):
                    yield cell.Text.createTextCursorByRange(cell)
            elif selections.supportsService(self.TextRangesNS):
                for tr in wrapUnoContainer(selections):
                    yield tr.Text.createTextCursorByRange(tr)

    def isOverlaping(self, rng1, rng2):
        if rng1.Text == rng2.Text:  # only deal with the same texts
            if rng1.Text.compareRegionStarts(rng1, rng2) < 0:
                rng1, rng2 = rng2, rng1
            # now we placed them in right order, lets compare end with start
            if rng1.Text.compareRegionStarts(rng1.End, rng2) < 0:
                return True
            return False

    def isInsideParagraph(self, rng=None):
        if rng is None:
            # get current selection
            selections = self.doc.getCurrentSelection()
            if selections is not None:
                if selections.supportsService(self.TableCursorNS):
                    return False
                if selections.supportsService(self.TextRangesNS):
                    rng = selections.getByIndex(0)
        if rng is not None:
            if self.isInsideCell(rng):  # cell is impossible to enum, bug
                return True
            enum = rng.createEnumeration()
            enum.nextElement()
            return not enum.hasMoreElements()
        return False

    def isTableSelected(self):
        """
        True if selection is TableCursor
        """
        selections = self.doc.getCurrentSelection()
        if selections is not None:
            if selections.supportsService(self.TableCursorNS):
                return True
        return False

    def isInsideCell(self, rng=None):
        if rng is None:
            selections = self.doc.getCurrentSelection()
            if selections is not None:
                if selections.supportsService(self.TableCursorNS):
                    return False  # selection is not IN cell
                elif selections.supportsService(self.TextRangesNS):
                    rng = selections.getByIndex(0)
        if rng is not None:
            return rng.Cell is not None
        return False

    def getSelectionEdges(self, selectionIndex=0):
        """
        returns (lhRange, rhRange)
        works well with texttablecursor
        """
        selections = self.doc.getCurrentSelection()
        if selections.supportsService(self.TableCursorNS):
            # inside the table
            vc = self.getViewCursor()
            if vc.TextTable is not None:
                cells = selections.RangeName.split(self.TableRangeNameSplitter)
                if len(cells) > 1:  # more than one cell selected
                    return (
                        vc.TextTable.getCellByName(cells[0]).Start,
                        vc.TextTable.getCellByName(cells[1]).End)
                else:
                    return(vc.TextTable.getCellByName(cells[0]).Start,
                           vc.TextTable.getCellByName(cells[0]).End)
        elif selections.supportsService(self.TextRangesNS):
            # we have a normal text selections
            if selections.Count > selectionIndex:
                sel = selections.getByIndex(selectionIndex)
                return (sel.Start, sel.End)

    def cursorFromSelection(self, selectionIndex=0):
        selectionIterator = self.iterateSelections()
        for cursorCount in range(selectionIndex + 1):
            try:
                cursor = next(selectionIterator)
            except StopIteration:
                pass
        return cursor

    def isAnythingSelected(self):
        selections = self.doc.getCurrentSelection()
        if selections.supportsService(self.TableCursorNS):
            return True  # Something is selected inside table
        elif selections.supportsService(self.TextRangesNS):
            if selections.Count > 1:
                return True
            if selections.Count > 0:
                selectionCursor = self.cursorFromSelection()
                if not selectionCursor.isCollapsed():
                    return True
        return False


class StyleUtilities(BaseUtilities):
    """
    Some short elementary routines for working on styles
    """

    def docHasParaStyle(self, paraStyleName):
        paraStyles = wrapUnoContainer(
            self.doc.StyleFamilies.getByName("ParagraphStyles"), "XName")
        return paraStyleName in paraStyles

    def docHasCharStyle(self, charStyleName):
        charStyles = wrapUnoContainer(
            self.doc.StyleFamilies.getByName("CharacterStyles"), "XName")
        return charStyleName in charStyles

    def docHasFont(self, fontName):
        """
        We need to find available fonts from current controller
        """
        currentWindow = self.doc.getCurrentController(
        ).getFrame().getContainerWindow()
        for f in currentWindow.getFontDescriptors():
            if f.Name == fontName:
                return True
        return False
