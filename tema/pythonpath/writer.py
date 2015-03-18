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
log = logging.getLogger('console_macro_debug')
import macrohelper
from pythonize import wrapUnoContainer

from com.sun.star.text.ControlCharacter import (  # noqa
                                                PARAGRAPH_BREAK,
                                                LINE_BREAK,
                                                HARD_HYPHEN,
                                                SOFT_HYPHEN,
                                                HARD_SPACE,
                                                APPEND_PARAGRAPH)

basic = None
current_doc = None

if 'XSCRIPTCONTEXT' in globals():
    basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa
    doc = basic.ThisComponent


def set_globals(ctx=None):
    global basic
    global current_doc
    if ctx is None and 'XSCRIPTCONTEXT' in globals():
        ctx = XSCRIPTCONTEXT  # noqa
    basic = macrohelper.StarBasicGlobals(ctx)
    current_doc = basic.ThisComponent


def current_macro():
    log.debug("Contact!")

    iu = IndexUtilities(basic.ThisComponent)
    iu.insertMark(iu.markKeysFromString(
        "Индексное вхождение:Первый уровень:Второй уровень"))
    msgl = []
    for im in iu.getMarks():
        msgl.append(iu.makeMarkPresentation(im))
    basic.MsgBox("\n".join(msgl))


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

    def __init__(self, *args, **kwargs):
        super(IndexUtilities, self).__init__(*args, **kwargs)
        self.Cursor = CursorUtilities(self.doc)
        self.Fields = FieldUtilities(self.doc)
        self.AlternativeTextList = []
        self.PrimaryKeyList = []
        self.SecondaryKeyList = []

    def iterMarks(self, rng):
        """
        iterate over indexMarks in range
        """
        for portion in self.Cursor.iterateTextPortions(rng):
            if portion.TextPortionType == "DocumentIndexMark":
                yield portion.DocumentIndexMark

    def keysFromString(self, markString):
        return {k: v for (k, v) in zip(
            self.MarkKeyNames,
            markString.split(self.MarkKeySeparator)
        )}

    def keysToString(self, markProperties):
        """
        reverse of indexMarkKeysFromString
        """
        markKeysList = []
        for k in self.MarkKeyNames:
            if k in markProperties:
                keyField = markProperties[k]
                if len(keyField):
                    markKeysList.append(keyField)
        return ":".join(markKeysList)

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
        lastMarkNum = len(self.getMarks())
        sel = self.Cursor.cursorFromSelection()
        if self.Cursor.isInsideParagraph(sel):
            markString = "%s<%s>" % (markString, lastMarkNum)
            markProperties = self.keysFromString(markString)
            self.insertMark(self.createMark(markProperties))
        else:
            # more than one paragraph selected
            sel = self.Cursor.forwards(sel)  # make left-to-right
            # Places 2 index marks at the start and end of textRange
            markString1 = "%s+<%s>" % (markString, lastMarkNum)
            markString2 = "%s=<%s>" % (markString, lastMarkNum+1)
            self.insertMark(self.createMark(
                self.keysFromString(markString1)),
                sel.getStart())
            self.insertMark(self.createMark(
                self.keysFromString(markString2)),
                sel.getEnd(),
                False)

    def stripMarkNumber(self, keyString):
        """
        remove =<123> or <123> or +<123> from the end of keyString
        """
        sepPosition = len(keyString)
        for sep in ('+', '=', '<'):
            if sep in keyString:
                sepPosition = min(sepPosition, keyString.find(sep))
        return keyString[:sepPosition]

    def cacheKeys(self):
        for m in self.getMarks():
            markProperties = Properties.dctFromProperties(m)
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
        marksRemoved = 0
        # if something is selected, look for all fields in selection
        if self.Cursor.isAnythingSelected():
            selectionCursor = self.Cursor.cursorFromSelection()
            for textField in self.Fields.iterateTextFields(selectionCursor):
                # check that it is our field
                # Dogtgtu
                log.debug("Checking that %s is in %s", self.ShowMarksVarName,
                          textField.Condition)
                if self.ShowMarksVarName in textField.Condition:
                    # find attached index mark
                    for mark in self.getAttachedIndexMarks(textField):
                        for m in self.getLinkedMarks(mark):
                            marksRemoved += 1
                            self.doc.Text.removeTextContent(m)
                    # remove presentation aswell
                    self.doc.Text.removeTextContent(textField)
        return marksRemoved

    def getAttachedIndexMarks(self, presetationField):
        cur = self.doc.Text.createTextCursorByRange(presetationField.Anchor)
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
        markString = self.keysToString(Properties.dctFromProperties(mark))
        if '=' in markString or '+' in markString:
            markNumber = self.readMarkNumber(markString)
            if markNumber is not None:
                return (mark, self.getMarkByNumber(markNumber + 1))
        # this is single mark
        return (mark,)

    def getMarkByNumber(self, markNumber):
        for m in self.getMarks():
            markProperties = Properties.dctFromProperties(m)
            markString = self.keysToString(markProperties)
            mn = self.readMarkNumber(markString)
            if mn == markNumber:
                return m

    def readMarkNumber(self, markString):
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
        self.doc.Text.insertTextContent(insertionPoint,
                                        mark, False)
        if givePresentation:
            field = self.Fields.createHiddenTextField(
                self.makeMarkPresentation(mark), varName=self.ShowMarksVarName)
            self.doc.Text.insertTextContent(insertionPoint,
                                            field, False)
            self.doc.TextFields.refresh()

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
        cur = self.doc.Text.createTextCursorByRange(rng)
        {"left": cur.collapseToStart, "right": cur.collapseToEnd}[edge]()
        return cur

    def forwards(self, rng):
        return self.createTextCursorByEdges(
            self.getEdge(rng), self.getEdge(rng, "right"))

    def backwards(self, rng):
        return self.createTextCursorByEdges(
            self.getEdge(rng, "right"), self.getEdge(rng))

    def iterateParagraphs(self, rng):
        return wrapUnoContainer(rng)

    def iterateTextPortions(self, rng):
        for para in self.iterateParagraphs(rng):
            for portion in wrapUnoContainer(para):
                if self.isOverlaping(portion, rng):
                    yield portion

    def isOverlaping(self, rng1, rng2):
        for i in range(1):
            rng1, rng2 = rng2, rng1
            if self.doc.Text.compareRegionStarts(
                    self.getEdge(rng1, "right"), self.getEdge(rng2)) < 1:
                return True
        return False

    def isInsideParagraph(self, rng):
        enum = rng.createEnumeration()
        enum.nextElement()
        return not enum.hasMoreElements()

    def cursorFromSelection(self, selectionIndex=0):
        selections = self.doc.getCurrentSelection()
        if selections.Count > selectionIndex:
            return self.doc.Text.createTextCursorByRange(
                selections.getByIndex(selectionIndex))

    def isAnythingSelected(self):
        selections = self.doc.getCurrentSelection()
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
        currentWindow = doc.getCurrentController(
        ).getFrame().getContainerWindow()
        for f in currentWindow.getFontDescriptors():
            if f.Name == fontName:
                return True
        return False
