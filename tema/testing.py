""" working with bookmarks"""

import logging
log = logging.getLogger(__name__)

import macrohelper
basic = None
doc = None

if 'XSCRIPTCONTEXT' in globals():
    basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa
    doc = basic.ThisComponent


def set_globals(ctx=None):
    global basic
    global doc
    if ctx is None and 'XSCRIPTCONTEXT' in globals():
        ctx = XSCRIPTCONTEXT  # noqa
    basic = macrohelper.StarBasicGlobals(ctx)
    doc = basic.ThisComponent


from pythonize import wrapUnoContainer, UnoDateConverter

from com.sun.star.text.ControlCharacter import (  # noqa
    PARAGRAPH_BREAK,
    LINE_BREAK,
    HARD_HYPHEN,
    SOFT_HYPHEN,
    HARD_SPACE,
    APPEND_PARAGRAPH)


class BaseUtilities:
    def __init__(self, _doc):
        self.doc = _doc


appendText = lambda t: doc.Text.insertString(
    doc.Text.getEnd(),
    t,
    False)


def appendPara(t):
    appendText(t)
    doc.Text.insertControlCharacter(doc.Text.getEnd(), PARAGRAPH_BREAK, False)


def current_macro():
    log.debug("Contact!")

    iu = IndexUtilities(basic.ThisComponent)
    iu.insertIndexMark(iu.indexMarkKeysFromString(
        "Индексное вхождение:Первый уровень:Второй уровень"))


class IndexUtilities(BaseUtilities):
    """
    Short routines for manipulating index marks in text
    """
    IndexMarkNS = "com.sun.star.text.DocumentIndexMark"
    IndexMarkKeySeparator = ":"

    def __init__(self, *args, **kwargs):
        super(IndexUtilities, self).__init__(*args, **kwargs)
        self.Cursor = CursorUtilities(self.doc)

    def showIndexMarks(self):
        """
        adds special field ShowIndexMarks with value 1 at the begining of doc
        """
        pass

    def indexMarkKeysFromString(self, indexMarkString):
        return {k: v for (k, v) in zip(
            ('AlternativeText', 'PrimaryKey', 'SecondaryKey'),
            indexMarkString.split(self.IndexMarkKeySeparator)
        )}

    def insertIndexMark(self, indexMarkProperties):
        """
        inserts index mark at the current cursor position
        sets properties from dct indexMarkProperties:
        - AlternativeText
        - PrimaryKey
        - SecondaryKey
        - IsMainEntry
        """
        indexMark = self.doc.createInstance(self.IndexMarkNS)
        Properties.setFromDict(indexMark, indexMarkProperties)
        self.doc.Text.insertTextContent(self.getInsertionPoint(),
                                        indexMark, False)

    def getInsertionPoint(self):
        """
        finds current cursor position to inser IndexMark to
        """
        return self.Cursor.getCurrentPosition()


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


class CursorUtilities(BaseUtilities):
    """
    Work with view and text cursors
    """

    def getViewCursor(self):
        return self.doc.CurrentController.getViewCursor()

    def getCurrentPosition(self):
        viewCursor = self.getViewCursor()
        return viewCursor.getStart()


class Properties:
    """
    Work with properties
    """

    @classmethod
    def setFromDict(cls, obj, dct):
        for k in dct:
            obj.setPropertyValue(k, dct[k])


def displayStyleProperties():
    paraStyles = doc.StyleFamilies.getByName("ParagraphStyles")
    for s in wrapUnoContainer(paraStyles, "XIndex"):
        if s.isInUse():
            appendPara("Style name: %s, display name: %s, properties:" % (
                s.Name, s.DisplayName))
            for p in s.getPropertySetInfo().getProperties():
                appendPara(p.Name)


def displayUsedParaStyles():
    dctFamilies = wrapUnoContainer(doc.StyleFamilies, "XName")
    paraStyles = dctFamilies["ParagraphStyles"]
    for s in wrapUnoContainer(paraStyles, "XIndex"):
        if s.isInUse():
            appendPara("%s" % s.DisplayName)


def displayAllStyles():
    dctFamilies = wrapUnoContainer(doc.StyleFamilies, "XName")
    log.debug("dctFamilies is %s", type(dctFamilies))
    for k in dctFamilies:
        appendPara("%s\t%s" % (dctFamilies[k].getCount(), k))
    # show styles inside families
    for familyName in dctFamilies:
        appendPara("%s" % familyName)
        i = 0
        log.debug("types of dctFamilies[%s] are: %s", familyName,
                  dctFamilies[familyName].Types)
        appendPara("No.\t Slug name\tName\tDisplay name")
        dctFamily = wrapUnoContainer(dctFamilies[familyName], "XName")
        for styleName in sorted(dctFamily):
            i += 1
            s = dctFamily[styleName]
            appendPara("%d:\t%s\t%s\t%s" % (
                i, styleName, s.Name, s.DisplayName))


def showProperties():
    propSetInfo = doc.getPropertySetInfo()
    props = propSetInfo.getProperties()
    for p in props:
        doc.Text.insertString(
            doc.Text.createTextCursor(), "%s\t%s\n" % (
                p.Name,
                doc.getPropertyValue(p.Name)), False)


def show_macro_name():
    vcursor = doc.CurrentController.ViewCursor
    doc.Text.insertString(vcursor.Start, "{}".format(__name__), False)


def addBookmark():
    """ just inserts simple bookmark"""
    doc = basic.ThisComponent
    cursor = doc.Text.createTextCursor()
    cursor.gotoStart(False)
    cursor.goRight(4, True)

    bookmark = doc.createInstance("com.sun.star.text.Bookmark")
    bookmark.Name = "Tema test bookmark"
    doc.Text.insertTextContent(cursor, bookmark, False)

""" working with character properties """


def changeFontRelief():
    paraCount = 0
    for para in wrapUnoContainer(doc.Text):
        paraCount += 1
        para.CharRelief = paraCount % 3


def removeFontRelief():
    from com.sun.star.text.FontRelief import NONE as NORELIEF
    paraCount = 0
    for para in wrapUnoContainer(doc.Text):
        paraCount += 1
        para.CharRelief = NORELIEF


"""Working with cursors:
    There are 2 type of cursors
    viewCursor = the real visible cursor, that knows all about current
    presentation of the text: pages, lines, screens ...
    textCursor = unvisible cursor that handles inner structure of the
    text, i.e. paragraphs, words, textparts...
"""


def InsertAtCursor(charcode=0xA9):
    """insert character at the current cursor position"""
    vcursor = doc.CurrentController.ViewCursor
    doc.Text.insertString(vcursor.Start, chr(charcode), False)


def countStatistics():
    """ count text statistics using text cursor"""
    cursor = doc.Text.createTextCursor()

    # count paragraphs
    cursor.gotoStart(False)
    paraCount = 0
    while cursor.gotoNextParagraph(False):
        paraCount += 1
    paraCount += 1

    # count sentences
    cursor.gotoStart(False)

"""
Learning to work with fields
"""

def DisplayFields():
    """ displays fields from current doc """
    fields = wrapUnoContainer(basic.ThisComponent.getTextFields())
    msg = ""
    for f in fields:
        msg += "{} = ".format(f.getPresentation(True))  # Field type
        if f.supportsService("com.sun.star.text.TextField.Annotation"):
            # for annotation we want to print author
            msg += "{} says {}".format(f.Author, f.Content)
        else:
            msg += f.getPresentation(False)  # String content
        msg += '\n'
    basic.MsgBox(msg, "Text fields")


def ShowFieldMasters():
    """ displays masterfields from current doc """

    fmasters = wrapUnoContainer(basic.ThisComponent.getTextFieldMasters())
    msg = ""
    for key in fmasters:
        field = fmasters[key]
        depCount = len(field.DependentTextFields)
        msg += "*** {} ***\n".format(key)
        msg += "{} master contains {} dependants\n".format(field.Name, depCount)
    basic.MsgBox(msg, "Text field masters")

fieldkey = lambda name: "com.sun.star.text.fieldmaster.User.{}".format(name)


def ProcessContract():
    """ first try to work with Art-tranzit contract"""
    fmasters = wrapUnoContainer(basic.ThisComponent.getTextFieldMasters())
    fk = fieldkey("DOCNUM")
    docnum = fmasters[fk]

    if docnum is not None:
        basic.MsgBox("DOCNUM found")
        docnum.Content = u"1453АБВABC"

    frame   = basic.ThisComponent.CurrentController.Frame
    dispatcher = basic.CreateUnoService("com.sun.star.frame.DispatchHelper")
    from com.sun.star.beans import PropertyValue
    dispatcher.executeDispatch(frame, ".uno:UpdateFields", "", 0, (PropertyValue(),))


def InsertTextField():
    """ inserts several text fields """

    from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK
    doc = basic.ThisComponent
    text = doc.Text
    doc.Text.insertControlCharacter(doc.Text.End, PARAGRAPH_BREAK, False)
    doc.Text.insertString(doc.Text.End, "The time is: ", False)
    dtField = doc.createInstance("com.sun.star.text.TextField.DateTime")
    dtField.IsFixed = True
    doc.Text.insertTextContent(text.End, dtField, False)

    doc.Text.insertControlCharacter(doc.Text.End, PARAGRAPH_BREAK, False)
    doc.Text.insertString(doc.Text.End, "А здесь текст по-русски, тоже из iPython. "
            "Питон позволяет склеивать предложения в абзацы просто перенося "
            "кавычки на следующую строку... Дальше будет аннотация.\n", False)
    anField = doc.createInstance("com.sun.star.text.TextField.Annotation")

    from datetime import timedelta
    somedayago = timedelta(days=-10)
    added = UnoDateConverter.today() + somedayago
    anField.Content = "Это аннотация на русском языке. Кодировка utf-8"
    anField.Author = "Tema"
    anField.Date = added.getUnoDate()
    doc.Text.insertTextContent(doc.Text.End, anField, False)


def InsertFieldMaster():
    doc = basic.ThisComponent
    text = doc.Text
    branch = "com.sun.star.text.FieldMaster.User"
    name = "TestField"
    fullname = "{}.{}".format(branch, name)

    fmasters = wrapUnoContainer(doc.getTextFieldMasters())
    if fullname in fmasters:
        dependants = fmasters[fullname].DependentTextFields
        if len(dependants) > 0:
            doc.Text.removeTextContent(dependants[0])
            basic.MsgBox("Removed one instance from doc")
        else:
            basic.MsgBox("No instances found, lets remove the master")
            fmasters[fullname].Content = ""
            fmasters[fullname].dispose()
    else:
        basic.MsgBox("Not found, lets create one")
        fmaster = doc.createInstance(branch)
        fmaster.Name = name
        fmaster.Content = "Hello from first inserted masterfield!"

        dependant = doc.createInstance("com.sun.star.text.TextField.User")
        dependant.attachTextFieldMaster(fmaster)
        doc.Text.insertTextContent(doc.Text.End, dependant, False)
        basic.MsgBox("One field inserted at the end...")


"""working with graphics"""


def insertImage():
    url = basic.InputBox("Please type image URL", "Path")
    if url:
        insertImageByURL(doc, url)


def insertImageByURL(doc, url):
    import unohelper
    cursor = doc.Text.createTextCursor()
    cursor.goToStart(False)
    img = doc.createInstance("com.sun.star.text.GraphicObject")
    img.GraphicURL = url
    img.Width = 6000
    img.Height = 8000
    img.AnchorType = unohelper.uno.getConstantByName(
            "com.sun.star.text.TextContentAnchorType.AS_CHARACTER")

    doc.Text.insertTextContent(cursor, img, False)


"""working with paragraphs"""


def enumerateParagraphs():
    line = "This doc contains {} paragraphs and {} tables"
    paraCount = tableCount = 0
    from com.sun.star.style.ParagraphAdjust import RIGHT
    for para in wrapUnoContainer(doc.Text):  # text has XEnumerateAccess interface
        log.debug("para is %s", para)
        log.debug(dir(para))

        if para.supportsService("com.sun.star.text.Paragraph"):
            paraCount += 1
            if not paraCount % 5:
                para.ParaAdjust = RIGHT

                # this is how we work with structures: we copy or instantiate
                # them and then assign to the property
                from com.sun.star.style import LineSpacing
                from com.sun.star.style.LineSpacingMode import LEADING
                # we could say linespacing = para.ParaLineSpacing
                linespacing = LineSpacing()
                linespacing.Mode = LEADING # just int constant
                linespacing.Height = 1000
                para.ParaLineSpacing = linespacing
        if para.supportsService("com.sun.star.text.Table"):
            tableCount += 1

    basic.MsgBox(line.format(paraCount, tableCount), "Statistics")


def enumerateTextSections():
    """Just as you can enumerate the paragraphs in a document,
    you can enumerate the text sections in a paragraph.
    Text within each enumerated portion uses the same
    properties and is of the same type."""

    paraCount = 0
    msg = ""
    for para in wrapUnoContainer(doc.Text):
        paraCount += 1
        msg += "{}. ".format(paraCount)
        for textpart in wrapUnoContainer(para):
            msg += "{}:".format(textpart.TextPortionType)
        msg += "\n"
        if not paraCount % 10:
            basic.MsgBox(msg, "Paragraph text sections")
            break
