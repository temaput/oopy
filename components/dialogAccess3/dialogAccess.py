"""
Python version of dialoghandler component from sdk examples
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
log = logging.getLogger(__name__)
log.addHandler(logging.StreamHandler())
log.setLevel(logging.DEBUG)

import uno  # noqa
import unohelper


from com.sun.star.awt import (XDialogEventHandler, XDialogProvider2)


class DialogAccessComponent3(unohelper.Base,
                             XDialogEventHandler, XDialogProvider2):

    """
    Implements DialogAccess component as in
    https://wiki.openoffice.org/wiki/Documentation/DevGuide/WritingUNO/
    Using_Dialogs_in_Components
    """
    implementationName = "org.openoffice.comp.pyuno.practica.DialogAccess3"
    dialogProviderNS = "com.sun.star.awt.DialogProvider2"
    ListControlsMethods = ("AlternativeTextListAction",
                           "PrimaryKeyListAction")

    def __init__(self, ctx):
        # store the component context for later use
        self.ctx = ctx
        self.dialog = None

    def createDialog(self, DialogURL):
        """
        presents dialog by accessing its path in library
        uses either model or frame (if no doc loaded)
        """

        smgr = self.ctx.getServiceManager()
        log.debug("Creating dialog with URL %s", DialogURL)
        log.debug("Context has attr ServiceManager? -%s",
                  hasattr(self.ctx, "ServiceManager"))

        obj = smgr.createInstanceWithContext(
            self.dialogProviderNS, self.ctx)

        self.dialog = obj.createDialogWithHandler(DialogURL, self)
        if self.dialog is not None:
            self.fillListBoxes()
            self.dialog.execute()

    def fillListBoxes(self):
        atl = self.dialog.getControl("AlternativeTextList")
        atl.addItems(
            self.getAlternativeTextList(), 0)
        pkl = self.dialog.getControl("PrimaryKeyList")
        pkl.addItems(
            self.getPrimaryKeyList(), 0)

    def getAlternativeTextList(self):
        return ("Первое вхождение", "Второе вхождение", "Третье вхождение")

    def getPrimaryKeyList(self):
        return ("Первый ключ", "Второй ключ", "Третий ключ")

    def callHandlerMethod(self, xDialog, EventObject, MethodName):
        """
        XDialogEventHandler universal event dispatcher
        """
        log.debug(
            "callHandlerMethod MethodName=%s",
            MethodName)
        log.debug(
            "callHandlerMethod called with ActionCommand=%s",
            EventObject.ActionCommand)
        log.debug("is method name in listcontrolmethods? %s",
                  MethodName in (self.ListControlsMethods))
        if MethodName in (self.ListControlsMethods):
            self.putOnKey(MethodName, EventObject.ActionCommand)
        elif MethodName == "Cancel":
            self.dialog.endExecute()
        return True

    def putOnKey(self, ListControl, line):
        log.debug("in putOnKey")
        if ListControl == self.ListControlsMethods[0]:
            self.cleanMarkString()
        else:
            line = ":%s" % line
        self.appendMarkString(line)

    def cleanMarkString(self):
        mstr = self.dialog.getControl("MarkString")
        mstr.Text = ""

    def appendMarkString(self, line):
        mstr = self.dialog.getControl("MarkString")
        mstr.Text = "%s%s" % (mstr.Text, line)


# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    DialogAccessComponent3,                        # UNO object class
    DialogAccessComponent3.implementationName,  # implementation name
    (DialogAccessComponent3.implementationName,),)
# list of implemented services
