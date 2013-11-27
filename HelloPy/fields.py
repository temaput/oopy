#set encoding=utf-8
"""
Learning to work with fields
"""

from tema.oo import macrohelper
basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)

from tema.oo.pythonize import wrapUnoContainer, UnoDateConverter


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
    text.insertControlCharacter(text.End, PARAGRAPH_BREAK, False)
    text.insertString(text.End, "The time is: ", False)
    dtField = doc.createInstance("com.sun.star.text.TextField.DateTime")
    dtField.IsFixed = True
    text.insertTextContent(text.End, dtField, False)

    text.insertControlCharacter(text.End, PARAGRAPH_BREAK, False)
    text.insertString(text.End, "А здесь текст по-русски, тоже из iPython. "
            "Питон позволяет склеивать предложения в абзацы просто перенося "
            "кавычки на следующую строку... Дальше будет аннотация.\n", False)
    anField = doc.createInstance("com.sun.star.text.TextField.Annotation")

    from datetime import timedelta
    somedayago = timedelta(days=-10)
    added = UnoDateConverter.today() + somedayago
    anField.Content = "Это аннотация на русском языке. Кодировка utf-8"
    anField.Author = "Tema"
    anField.Date = added.getUnoDate()
    text.insertTextContent(text.End, anField, False)

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
            text.removeTextContent(dependants[0])
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
        text.insertTextContent(text.End, dependant, False)
        basic.MsgBox("One field inserted at the end...")
