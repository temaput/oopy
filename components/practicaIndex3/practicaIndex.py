"""
Main component for working with indexes in Practica
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
log = logging.getLogger("pyuno")
log.addHandler(logging.NullHandler())
log.setLevel(logging.WARNING)


import uno  # noqa
import unohelper


from macrohelper import StarBasicGlobals
import writer
from comphelper import ProtocolHandlerComponentHelper
from comphelper import DialogAccessComponentHelper


class PracticaIndex(DialogAccessComponentHelper,
                    ProtocolHandlerComponentHelper):

    ImplementationName = "org.openoffice.comp.pyuno.practica.Index"
    InsertionDialogName = "vnd.sun.star.script:libPracticaIndexBasic.IndexMarkerInsertDialog?location=application"  # noqa

    def __init__(self, ctx, *args, **kwargs):
        super(PracticaIndex, self).__init__(ctx, *args, **kwargs)
        self.Basic = StarBasicGlobals(self.ctx)
        self.doc = self.Basic.ThisComponent
        self.iu = writer.IndexUtilities2(self.doc)
        self.cu = writer.CursorUtilities(self.doc)
        self.smgr = self.ctx.getServiceManager()

    def prepareDialog(self):
        if self.iu.LastMarkNum is None:
            self.iu.rebuildCache(self.doc)
        self.fillListBoxes()
        self.fillMarkString()

    def fillMarkString(self):
        """
        take the selection and put it to MarkString
        """
        log.debug("preparing MarkString")
        cur = self.cu.cursorFromSelection()
        if not self.cu.isInsideParagraph(cur) or cur.isCollapsed():
            cur.collapseToStart()
            cur.gotoEndOfSentence(True)

        log.debug("markSTring = %s", cur.String)
        self.appendMarkString(cur.String)

    def fillListBoxes(self):
        atl = self.dialog.getControl("AlternativeTextList")
        atl.addItems(
            self.getAlternativeTextList(), 0)
        pkl = self.dialog.getControl("PrimaryKeyList")
        pkl.addItems(
            self.getPrimaryKeyList(), 0)

    def getAlternativeTextList(self):
        return tuple(self.iu.FirstEntryList)

    def getPrimaryKeyList(self):
        return tuple(self.iu.SecondEntryList)

    def cleanMarkString(self):
        mstr = self.dialog.getControl("MarkString")
        mstr.Text = ""

    def appendMarkString(self, line):
        mstr = self.dialog.getControl("MarkString")
        mstr.Text = "%s%s" % (mstr.Text, line)

    def getMarkString(self):
        mstr = self.dialog.getControl("MarkString")
        return mstr.Text

    def insertMarker(self):
        MarkString = self.getMarkString()
        log.debug("inserting mark string %s", MarkString)
        self.iu.makeMarkHere(MarkString)

    def AlternativeTextListAction(self, xDialog, EventObject, MethodName):
        self.cleanMarkString()
        line = EventObject.ActionCommand
        self.appendMarkString(line)
        return True

    def PrimaryKeyListAction(self, xDialog, EventObject, MethodName):
        line = ":%s" % EventObject.ActionCommand
        self.appendMarkString(line)
        return True

    def Cancel(self, *args):
        self.dialog.endExecute()
        return True

    def OK(self, *args):
        self.insertMarker()
        log.debug("Marker inserted (in OK)")
        self.dialog.endExecute()
        return True

    def IndexMarkInsertDispatch(self):
        """
        dispatch command to show form for inserting index mark
        """
        try:
            self.createDialog(self.InsertionDialogName)
        except writer.BadSelection:
            self.Basic.MsgBox("Try another selection!")
        # me = self.smgr.createInstance(self.ImplementationName)
        # if me is not None:
        #     me.createDialog(self.InsertionDialogName)

    def IndexMarkRemoveDispatch(self):
        try:
            self.iu.removeMarkHere()
        except writer.BadSelection:
            self.Basic.MsgBox("Try another selection!")

    def ToggleMarkPresentationsDispatch(self):
        self.iu.toggleMarks()

    def CountIndexMarks(self):
        di = self.doc.createInstance("com.sun.star.text.DocumentIndex")
        marks = di.DocumentIndexMarks
        self.Basic.MsgBox("%s" % len(marks))


# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
ComponentClass = PracticaIndex

g_ImplementationHelper.addImplementation(
    ComponentClass,                        # UNO object class
    ComponentClass.ImplementationName,  # implementation name
    (ComponentClass.ImplementationName,),)
# list of implemented services
