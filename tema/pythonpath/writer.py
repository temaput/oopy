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

import re

from pythonize import wrapUnoContainer

from com.sun.star.text.ControlCharacter import (  # noqa
                                                PARAGRAPH_BREAK,
                                                LINE_BREAK,
                                                HARD_HYPHEN,
                                                SOFT_HYPHEN,
                                                HARD_SPACE,
                                                APPEND_PARAGRAPH)


from macrohelper import colors

from utils import Bunch


class BadSelection(ValueError):
    pass


class BaseUtilities:
    def __init__(self, _doc):
        self.doc = _doc


class Properties:
    """
    Work with properties
    """

    def __init__(self, obj):
        psi = obj.getPropertySetInfo()
        object.__setattr__(self, "psi", psi)
        object.__setattr__(self, "obj", obj)

    def __setattr__(self, key, value):
        if self.psi.hasPropertyByName(key):
            self.obj.setPropertyValue(key, value)
        else:
            raise AttributeError("No such property %s" % key)

    def __getattr__(self, key):
        if self.psi.hasPropertyByName(key):
            return self.obj.getPropertyValue(key)
        else:
            raise AttributeError("No such property %s" % key)

    def __setitem__(self, key, value):
        setattr(self, key, value)

    def __getitem__(self, key):
        return getattr(self, key)

    @staticmethod
    def setFromDict(obj, dct):
        for k in dct:
            obj.setPropertyValue(k, dct[k])

    @staticmethod
    def propTupleFromDict(propDct):
        """
        Returns tuple of PropertyValue elements for setting as Attributes
        """
        from com.sun.star.beans import PropertyValue
        propList = []
        for k in propDct:
            pv = PropertyValue()
            pv.Name = k
            pv.Value = propDct[k]
            propList.append(pv)
        return tuple(propList)

    @staticmethod
    def dctFromProperties(obj, property_names=()):
        PropertySetInfo = obj.getPropertySetInfo()
        dct = {}
        for p in PropertySetInfo.Properties:
            if property_names and p.Name not in property_names:
                continue
            dct[p.Name] = obj.getPropertyValue(p.Name)
        return dct


class TextUtilities(BaseUtilities):
    """
    Work with text, paragraphs
    """

    def insertTextAtRange(self, rng, t):
        """
        Inserts string at the start of rng
        returns TextCursor enclosing inserted text
        """
        cur = rng.Text.createTextCursorByRange(rng)
        cur.Text.insertString(
            cur,
            t,
            True)
        return cur

    def insertParagraph(self, rng, t):
        """
        Inserts text and paragraph break at the Start of rng
        returns textCursor enclosing inseted text
        """
        cur = self.insertTextAtRange(rng, t)
        cur.Text.insertControlCharacter(
            cur.End, PARAGRAPH_BREAK, False)
        return cur

    def appendText(self, t):
        self.doc.Text.insertString(
            self.doc.Text.getEnd(),
            t,
            False)

    def markRange(self, rng, color=None):
        if color is None:
            color = colors.yellow
        rng.CharBackColor = color

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
    MarkPresentationMask = "<XE %s>"
    ShowMarksVarName = "showIndexMarks"
    MarkNumberBrackets = ("{", "}")
    MarkPropertiesTemplate = {"AlternativeText": "",
                              "PrimaryKey": "",
                              "SecondaryKey": ""}

    signs = Bunch(
        diapasonOpening="+",
        diapasonClosing="="
    )
    MaxLevels = 3  # maximal index depth
    # ========================================
    # Class based cache
    # ========================================
    RefCount = 0
    FirstEntryList = []
    SecondEntryList = []
    ThirdEntryList = []
    LastMarkNum = None
    _doc = None
    MarkCacheDict = {}

    def __del__(self):
        log.debug("Remove instance of IndexUtilities")

    def __init__(self, *args, **kwargs):
        super(IndexUtilities, self).__init__(*args, **kwargs)
        self.Cursor = CursorUtilities(self.doc)
        self.Fields = FieldUtilities(self.doc)
        self.incrementRefCount()
        log.debug("initializing %s instance of IndexUtitlities", self.RefCount)
        if self._doc is not None and self.doc != self._doc:
            self.resetCache()

    @classmethod
    def incrementRefCount(cls):
        cls.RefCount += 1

    def iterPresentationFields(self, rng=None):
        for tf in self.Fields.iterateTextFields(rng):
            if self.Fields.isHiddenTextField(tf) and\
                    self.ShowMarksVarName in tf.Condition:
                yield tf

    def iterMarks(self, rng):
        """
        iterate over indexMarks in range
        """
        for portion in self.Cursor.iterateTextPortions(rng):
            if portion.TextPortionType == "DocumentIndexMark":
                yield portion.DocumentIndexMark

    def killPresentationFields(self):
        fields_killed = 0
        for tf in self.iterPresentationFields():
            fields_killed += 1
            tf.dispose()
        return fields_killed

    def rebuildPresentationFields(self):
        self.killPresentationFields()
        processed_marks = []
        for m in self.getMarks():
            mark = self.getLinkedMarks(m)[0]
            if mark not in processed_marks:
                self.givePresentation(mark)
                processed_marks.append(mark)
        return len(processed_marks)

    def convert_old_index_markers(self):
        """
        convert bad im entries (index number like <111>) to
        good (index number like {111}
        """
        for im in self.getMarks():
            at = im.AlternativeText
            im.AlternativeText = at.translate(str.maketrans("<>", "{}"))

    @staticmethod
    def keysFromString(markString):
        markProperties = IndexUtilities.MarkPropertiesTemplate.copy()
        markProperties.update({k: v for (k, v) in zip(
            IndexUtilities.MarkKeyNames,
            markString.split(IndexUtilities.MarkKeySeparator)
        )})
        return markProperties

    @staticmethod
    def keysToList(markProperties):
        markKeysList = []
        for k in IndexUtilities.MarkKeyNames:
            if k in markProperties:
                keyField = markProperties[k]
                if len(keyField):
                    markKeysList.append(keyField)
        return markKeysList

    @classmethod
    def keysToString(cls, markProperties):
        """
        reverse of indexMarkKeysFromString
        """
        return cls.MarkKeySeparator.join(cls.keysToList(markProperties))

    def makeMarkPresentation(self, mark, mask=None):
        """
        Prepares mark presentation for hidden text field
        aka {"XE" Name:First key:SecondaryKey}
        """
        mask = mask or self.MarkPresentationMask
        markProperties = Properties.dctFromProperties(mark)
        return mask % self.keysToString(markProperties)

    def printIndex(self, targetDoc=None):
        """
        Collects all markEntries and creates Index tree from them
        """
        from indexmaker import IndexMaker
        from io import StringIO
        sio = StringIO()
        presentationMask = "%s"
        vcur = self.Cursor.getViewCursor()
        log.info("Printing index from document")
        marks = self.getMarks()
        log.debug("Marks len found: %s", len(marks))
        for im in marks:
            # log.debug(im)
            imtext = self.makeMarkPresentation(im, presentationMask)
            log.debug("imtext: %s", imtext)
            vcur.gotoRange(im.Anchor, False)
            page = vcur.Page
            log.debug("index: %s\t%s",  (imtext, page))
            print("%s\t%s" % (imtext, page), file=sio)
        sio.seek(0)
        im = IndexMaker()
        im.makeIndex(sio, targetDoc)

    def makeMarkHere(self, markString):
        """
        Check for selection if it lies acros more than 1 paragraph mark range
        """
        if self.LastMarkNum is None:
            self.rebuildCache(self.doc)
        lastMarkNum = self.LastMarkNum
        if self.Cursor.isInsideParagraph() or self.Cursor.isInsideCell():
            markString = "%s{%s}" % (markString, lastMarkNum)
            markProperties = self.keysFromString(markString)
            mark = self.createMark(markProperties)
            self.insertMark(mark)
            log.debug("mark inserted, LMN=%s", self.LastMarkNum)
            self.addMarkToCache(mark, lastMarkNum)
            self.addMarkKeysToCache(markProperties)
            log.debug("cache updated")
            self.incrementLastMarkNum()
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
                self.keysFromString("%s+{%s}" % (markString, lastMarkNum)))
            mark2 = self.createMark(
                self.keysFromString("%s={%s}" % (markString, lastMarkNum + 1)))
            self.insertMark(mark1, selEdges[0])
            self.insertMark(mark2, selEdges[1], False)
            self.addMarkToCache(mark1, lastMarkNum)
            self.addMarkToCache(mark2, lastMarkNum + 1)
            self.incrementLastMarkNum(2)
        self.doc.TextFields.refresh()

    @classmethod
    def incrementLastMarkNum(cls, val=1):
        cls.LastMarkNum += val

    @staticmethod
    def stripMarkNumber(keyString):
        """
        remove ={123} or {123} or +{123} from the end of keyString
        """
        sepPosition = len(keyString)
        for sep in ('+', '=', '{'):
            if sep in keyString:
                sepPosition = min(sepPosition, keyString.find(sep))
        return keyString[:sepPosition]

    @classmethod
    def removeMarkFromCache(cls, key):
        cls.MarkCacheDict[key] = None

    @classmethod
    def addMarkToCache(cls, mark, markNumber):
        log.debug("add mark number %s to cache", markNumber)
        cls.MarkCacheDict[markNumber] = mark

    @classmethod
    def resetCache(cls):
        log.debug("reseting cache")
        cls.FirstEntryList = []
        cls.SecondEntryList = []
        cls.ThirdEntryList = []
        cls.MarkCacheDict = {}
        cls.LastMarkNum = None
        cls._doc = None

    @classmethod
    def rebuildCache(cls, doc):
        DocumentIndex = doc.createInstance(cls.IndexNS)
        marks = DocumentIndex.DocumentIndexMarks
        cls.LastMarkNum = len(marks)
        cls._doc = doc
        for m in marks:
            markProperties = Properties.dctFromProperties(m)
            cls.addMarkKeysToCache(markProperties)
            markKeysList = cls.keysToList(markProperties)
            markNumber = cls.readMarkNumber(markKeysList)
            if markNumber is not None:
                cls.LastMarkNum = max(cls.LastMarkNum, markNumber)
                cls.addMarkToCache(m, markNumber)

    @classmethod
    def addMarkKeysToCache(cls, markProperties):
        markKeysList = cls.keysToList(markProperties)
        log.debug("adding mark keys to Cache: %s", markKeysList)
        lastEntry = markKeysList[-1]
        markKeysList[-1] = cls.stripMarkNumber(lastEntry)

        if markKeysList[0] not in cls.FirstEntryList:
            cls.FirstEntryList.append(markKeysList[0])
        if len(markKeysList) > 1 \
                and markKeysList[1] not in cls.SecondEntryList:
            cls.SecondEntryList.append(markKeysList[1])
        if len(markKeysList) > 2 \
                and markKeysList[2] not in cls.ThirdEntryList:
            cls.ThirdEntryList.append(markKeysList[2])

        log.debug("FirstEntryList len = %s", len(cls.FirstEntryList))
        log.debug("SecondEntryList len = %s", len(cls.SecondEntryList))

    def removeMarkHere(self):
        """
        Check current selection for any presentations
        and get rid of them and their marks
        """

        if self.LastMarkNum is None:
            self.rebuildCache(self.doc)
        marksRemoved = 0
        # if something is selected, look for all fields in selection
        if self.Cursor.isAnythingSelected():
            for selectionCursor in self.Cursor.iterateSelections():
                for textField in self.iterPresentationFields(selectionCursor):
                    # check that it is our field
                    log.debug("Checking that %s is in %s",
                              self.ShowMarksVarName,
                              textField.Condition)
                    # find attached index mark
                    markKeys = self.parsePresentationField(textField)
                    if markKeys is not None:
                        markNumber = self.readMarkNumber(
                            self.keysToList(markKeys))
                        mark = self.getMarkByNumber(markNumber)
                        if mark is not None:
                            for m in self.getLinkedMarks(mark):
                                marksRemoved += 1
                                m.dispose()
                            self.removeMarkFromCache(markNumber)
                            if marksRemoved > 1:
                                self.removeMarkFromCache(markNumber + 1)
                    # remove presentation aswell
                    textField.dispose()
        return marksRemoved

    def getAttachedIndexMark(self, presentationField):
        markNumber = None
        markKeys = self.parsePresentationField(presentationField)
        if markKeys is not None:
            markNumber = self.readMarkNumber(markKeys)
        cur = self.Cursor.createTextCursorByANYRage(
            presentationField.Anchor)
        cur.gotoEndOfSentence(True)
        nearbyMarks = tuple(self.iterMarks(cur))
        if nearbyMarks and markNumber is None:
            return nearbyMarks[0]
        else:
            for mark in nearbyMarks:
                if markNumber == self.getMarkNumber(mark):
                    return mark

    def getAttachedPresentationField(self, mark):
        """
        travel cursor backwards to find the closest presentation field
        """
        mark = self.getLinkedMarks(mark)[0]  # always look for first mark
        markNumber = self.getMarkNumber(mark)
        cur = self.Cursor.createTextCursorByANYRage(mark.Anchor)
        cur.gotoStartOfSentence(True)
        for tf in self.Fields.iterateTextFields(cur):
            _markKeys = self.parsePresentationField(tf)
            if _markKeys is not None:
                _markNumber = self.readMarkNumber(_markKeys)
                if markNumber == _markNumber:
                    return tf

    def getMarkKeys(self, mark):
        return Properties.dctFromProperties(mark, self.MarkKeyNames)

    def getMarkNumber(self, mark):
        return self.readMarkNumber(self.keysToList(self.getMarkKeys(mark)))

    def getLinkedMarks(self, mark):
        """
        Inspects if it is a diapason marker
        returns tuple with mark and its opposite (if exist)
        """
        markKeysList = self.keysToList(self.getMarkKeys(mark))
        markString = markKeysList[-1]  # we only check the last element
        if '=' in markString or '+' in markString:
            markNumber = self.readMarkNumber(markKeysList)
            if markNumber is not None:
                oppositeMarkNumber = markNumber + 1 if '+' in markString \
                    else markNumber - 1
                if oppositeMarkNumber > 0:
                    mark2 = self.getMarkByNumber(oppositeMarkNumber)
                    if mark2 is not None:
                        if markNumber < oppositeMarkNumber:
                            return (mark, mark2)
                        else:
                            return (mark2, mark)
        # this is single mark
        return (mark,)

    def getMarkByNumber(self, markNumber):
        log.debug("is mark number %s in cache? %s",
                  markNumber, markNumber in self.MarkCacheDict)
        if markNumber in self.MarkCacheDict:
            return self.MarkCacheDict[markNumber]

    def parsePresentationField(self, presentationField, presentationMask=None):
        presentationMask = presentationMask or self.MarkPresentationMask
        regexp = presentationMask.replace("%s", "(?P<markString>.*)")
        m = re.search(regexp, presentationField.Anchor.String)
        if m is not None:
            markString = m.group("markString")
            if markString is not None:
                return self.keysFromString(markString)

    @staticmethod
    def readMarkNumber(markKeysList):
        markString = markKeysList[-1]  # we only look in last entry
        bracket1 = markString.find(IndexUtilities.MarkNumberBrackets[0]) + 1
        bracket2 = markString.find(IndexUtilities.MarkNumberBrackets[1])
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
            self.givePresentation(mark, cur)

    def givePresentation(self, mark, cur=None):
            field = self.Fields.createHiddenTextField(
                self.makeMarkPresentation(mark), varName=self.ShowMarksVarName)

            if cur is None:
                cur = self.Cursor.createTextCursorByANYRage(
                    mark.Anchor)
            cur.goLeft(1, False)
            cur.Text.insertTextContent(cur, field, False)

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

    @classmethod
    def keysFromString(cls, markString):
        markTuple = [k.strip() for k in
                     markString.split(cls.MarkKeySeparator)]
        markKeys = cls.MarkPropertiesTemplate.copy()
        markKeys['AlternativeText'] = markTuple[-1]
        if len(markTuple) > 1:
            markKeys['PrimaryKey'] = markTuple[0]
            if len(markTuple) > 2:
                markKeys['SecondaryKey'] = markTuple[1]
        return markKeys

    @staticmethod
    def keysToList(markProperties):
        markKeysList = [markProperties['AlternativeText']]
        if len(markProperties['PrimaryKey']):
            markKeysList.insert(0, markProperties['PrimaryKey'])
            if len(markProperties['SecondaryKey']):
                markKeysList.insert(1, markProperties['SecondaryKey'])
        return markKeysList


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

    def iterateTextFields(self, rng=None):
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

    def isTypeOfField(self, textField, textFieldType):
        branch = "%s.%s" % (self.TextFieldNS, textFieldType)
        return textField.supportsService(branch)

    def isHiddenTextField(self, textField):
        return self.isTypeOfField(textField, "HiddenText")

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
    TextTableNS = "com.sun.star.text.TextTable"
    TableRangeNameSplitter = ":"

    def getViewCursor(self):
        return self.doc.CurrentController.getViewCursor()

    def getCurrentPosition(self):
        viewCursor = self.getViewCursor()
        return viewCursor.getStart()

    """
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
    """

    def createTextCursorByANYRage(self, rng):
        if rng.Cell is not None:
            return rng.Cell.Text.createTextCursorByRange(rng)
        return rng.Text.createTextCursorByRange(rng)

    def iterateTableCells(self, cellRangeName=None, tbl=None):
        """
        gets table obj and iterates within its cells
        """
        opened = False
        closed = False
        cellRange = None
        if tbl is None:
            # no table given: get it from selection
            selections = self.doc.getCurrentSelection()
            if selections is not None:
                if selections.supportsService(self.TableCursorNS):
                    # several cells selected
                    vc = self.getViewCursor()
                    tbl = vc.TextTable
                    if cellRangeName is None:
                        cellRangeName = selections.RangeName
                    for cell in self.iterateTableCells(cellRangeName, tbl):
                        yield cell
                elif selections.supportsService(self.TextRangesNS):
                    # possibly cursor is inside the table
                    sel = selections.getByIndex(0)
                    tbl = sel.TextTable
                    if sel.Cell is not None:
                        yield sel.Cell
        else:
            # tbl is given
            if cellRangeName is not None:
                cellRange = cellRangeName.split(self.TableRangeNameSplitter)
                if len(cellRange) == 1:
                    yield tbl.getCellByName(cellRange[0])
                    closed = True
            else:
                # no cellRange: iterate whole table
                opened = True

            for cellName in tbl.getCellNames():
                if closed:
                    break
                if cellRange and cellName == cellRange[0]:
                    opened = True
                if cellRange and cellName == cellRange[1]:
                    closed = True
                if opened:
                    yield tbl.getCellByName(cellName)

    def iterateParagraphs(self, rng=None):
        if rng is None:
            # no range given iterate whole text
            rng = self.doc.Text
        for para in wrapUnoContainer(rng):
            if para.supportsService(self.TextTableNS):
                # in table
                for cell in self.iterateTableCells(tbl=para):
                    for cellpara in wrapUnoContainer(cell.Text):
                        yield cellpara
            else:
                yield para

    def iterateTableTextPortions(self, tbl):
        for cell in self.iterateTableCells(tbl=tbl):
            for para in wrapUnoContainer(cell.Text):
                for portion in wrapUnoContainer(para):
                    yield portion

    def iterateTextPortions(self, rng=None):
        if rng is None:
            enclosingRng = self.doc.Text
        else:
            if self.isInsideCell(rng):
                enclosingRng = rng.Cell
            else:
                enclosingRng = rng

        for para in self.iterateParagraphs(enclosingRng):
            for portion in wrapUnoContainer(para):
                if rng is None or self.isOverlaping(portion, rng):
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

    @staticmethod
    def isOverlaping(rng1, rng2):
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
    ParagraphStyleNS = "com.sun.star.style.ParagraphStyle"
    DefaultParaStyleName = "Standard"

    def docHasParaStyle(self, paraStyleName):
        paraStyles = wrapUnoContainer(
            self.doc.StyleFamilies.getByName("ParagraphStyles"), "XName")
        return paraStyleName in paraStyles

    def createParaStyle(self, paraStyleName, styleProperties,
                        ParentStyle=None):
        family = wrapUnoContainer(
            self.doc.StyleFamilies.getByName("ParagraphStyles"), "XName")
        if ParentStyle is None or ParentStyle not in family:
            ParentStyle = self.DefaultParaStyleName
        if "FollowStyle" in styleProperties:
            if styleProperties["FollowStyle"] not in family:
                styleProperties["FollowStyle"] = self.DefaultParaStyleName
        if paraStyleName not in family:
            newStyle = self.doc.createInstance(self.ParagraphStyleNS)
            newStyle.setParentStyle(ParentStyle)
            Properties.setFromDict(newStyle, styleProperties)
            family[paraStyleName] = newStyle

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


class FindReplaceUtilities:
    """
    Provides find/replace iteration routines
    boolean     SearchBackwards If TRUE, the search is done backwards in the
document. More...

    boolean     SearchCaseSensitive If TRUE, the case of the letters is
important for the match. More...

    boolean     SearchWords If TRUE, only complete words will be found. More...

    boolean     SearchRegularExpression If TRUE, the search string is evaluated
as a regular expression. More...

    boolean     SearchStyles If TRUE, it is searched for positions where the
paragraph style with the name of the search pattern is applied. More...

    boolean     SearchSimilarity If TRUE, a "similarity search" is performed.
More...

    boolean     SearchSimilarityRelax If TRUE, all similarity rules are applied
together. More...

    short   SearchSimilarityRemove This property specifies the number of
characters that may be ignored to match the search pattern. More...

    short   SearchSimilarityAdd specifies the number of characters that must be
added to match the search pattern. More...

    short   SearchSimilarityExchange This property specifies the number of
characters that must be replaced to match the search pattern. More...
    """

    def __init__(self, doc, *args, **kwargs):
        object.__setattr__(self, 'doc', doc)
        self.createDescriptor()
        if args:
            self.descriptor.update(args[0])

    def __setattr__(self, key, value):
        setattr(self.descriptor, key, value)

    def __getattr__(self, key):
        return getattr(self.descriptor, key)

    def createDescriptor(self):
        descriptor = Properties(self.doc.createSearchDescriptor())
        object.__setattr__(self, 'descriptor', descriptor)

    def _iterFromStart(self, descriptor):
        found = self.doc.findFirst(descriptor)
        while(found is not None):
            yield found
            found = self.doc.findNext(found.End, descriptor)

    def _iterInRange(self, descriptor, rng):
        curStart = rng.Text.createTextCursorByRange(rng)
        curStart.collapseToStart()
        found = self.doc.findNext(curStart, descriptor)
        while(found is not None):
            if CursorUtilities.isOverlaping(found, rng):
                yield found
                found = self.doc.findNext(found.End, descriptor)
            else:
                break

    def __call__(self, SearchString, ReplaceString=None, **kwargs):
        descriptor = self.descriptor.obj
        descriptor.SearchString = SearchString
        if ReplaceString is not None:
            descriptor.ReplaceString = ReplaceString
            return self.doc.replaceAll(descriptor)
        elif 'searchAll' in kwargs:
            return wrapUnoContainer(self.doc.findAll(descriptor))
        elif 'searchRange' in kwargs:
            return self._iterInRange(descriptor, kwargs['searchRange'])
        else:
            return self._iterFromStart(descriptor)

    def setSearchAttributes(self, attrDct):
        """ Set additional search attributes as dict
            can go with CharacterProperties and ParagraphProperties
        """
        self.descriptor.obj.setSearchAttributes(
            Properties.propTupleFromDict(attrDct))

    def setReplaceAttributes(self, attrDct):
        """ Set additional replace attributes as dict
            can go with CharacterProperties and ParagraphProperties
        """
        self.descriptor.obj.setReplaceAttributes(
            Properties.propTupleFromDict(attrDct))


class DocumentUtilities:
    """
    creating documents
    """
    NewDocumentNS = "private:factory/swriter"

    def __init__(self, desktop):
        self.desktop = desktop

    def createDocument(self, docName=None):
        return self.desktop.loadComponentFromURL(
            self.NewDocumentNS, "_blank", 0, ())


class BookmarkUtilities(BaseUtilities):

    def getBookmarksDict(self):
        return wrapUnoContainer(self.doc.Bookmarks,
                                "XNameAccess")
