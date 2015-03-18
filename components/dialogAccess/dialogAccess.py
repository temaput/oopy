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


class DialogAccessComponent(unohelper.Base,
                            XDialogEventHandler, XDialogProvider2):

    """
    Implements DialogAccess component as in
    https://wiki.openoffice.org/wiki/Documentation/DevGuide/WritingUNO/
    Using_Dialogs_in_Components
    """
    implementationName = "org.openoffice.comp.pyuno.demo.DialogAccess"
    dialogProviderNS = "com.sun.star.awt.DialogProvider2"
    desktopNS = "com.sun.star.frame.Desktop"

    def __init__(self, ctx):
        # store the component context for later use
        self.ctx = ctx

    def createDialog(self, DialogURL):
        """
        presents dialog by accessing its path in library
        uses either model or frame (if no doc loaded)
        """

        smgr = self.ctx.getServiceManager()
        log.debug("Creating dialog with URL %s", DialogURL)
        log.debug("Context is %s", self.ctx)
        log.debug("Context has attr getDesktop? -%s", hasattr(self.ctx,
                                                              "getDesktop"))
        log.debug("Context has attr ServiceManager? -%s",
                  hasattr(self.ctx, "ServiceManager"))

        desktop = smgr.createInstanceWithContext(
            self.desktopNS, self.ctx)
        xModel = desktop.getCurrentComponent()
        if xModel is not None:
            obj = smgr.createInstanceWithArgumentsAndContext(
                self.dialogProviderNS, (xModel,), self.ctx)
        else:
            obj = smgr.createInstanceWithContext(
                self.dialogProviderNS, self.ctx)

        xDialog = obj.createDialogWithHandler(DialogURL, self)
        log.debug("xDialog is %s", xDialog)
        if xDialog is not None:
            xDialog.execute()

    def callHandlerMethod(self, xDialog, EventObject, MethodName):
        """
        XDialogEventHandler universal event dispatcher
        """
        log.debug(
            "callHandlerMethod called with EventObject=%s, MethodName=%s",
            EventObject, MethodName)
        return True


# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    DialogAccessComponent,                        # UNO object class
    DialogAccessComponent.implementationName,  # implementation name
    (DialogAccessComponent.dialogProviderNS,),)  # list of implemented services
