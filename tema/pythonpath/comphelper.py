"""
Classes to help build my components aka *.oxt
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
log = logging.getLogger("pyuno.helpers")

import uno  # noqa
import unohelper


from com.sun.star.awt import (XDialogEventHandler, XDialogProvider2)


class DialogAccessComponentHelper(unohelper.Base,
                                  XDialogEventHandler, XDialogProvider2):

    """
    Implements DialogAccess component as in
    https://wiki.openoffice.org/wiki/Documentation/DevGuide/WritingUNO/
    Using_Dialogs_in_Components
    """
    DialogProviderNS = "com.sun.star.awt.DialogProvider2"

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
            self.DialogProviderNS, self.ctx)

        self.dialog = obj.createDialogWithHandler(DialogURL, self)
        if self.dialog is not None:
            self.prepareDialog()
            self.dialog.execute()

    def prepareDialog(self):
        pass

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
        log.debug("Do we have attr %s? %s",
                  MethodName, hasattr(self, MethodName))
        if hasattr(self, MethodName):
            return getattr(self, MethodName)(xDialog, EventObject, MethodName)


from com.sun.star.frame import (XDispatchProvider,
                                XDispatch)


class ProtocolHandlerComponentHelper(unohelper.Base,
                                     XDispatchProvider, XDispatch):
    """
    Implements ProtocolHandler
    component as in https://wiki.openoffice.org/wiki/Documentation/DevGuide/
    WritingUNO/Implementation
    """
    ImplementationName = None

    def __init__(self, ctx):
        # store the component context for later use
        self.ctx = ctx

    def queryDispatch(self, aURL, sTargetFrameName, iSearchFlags):
        log.debug("looking for %s in %s", self.ImplementationName,
                  aURL.Protocol)
        if self.ImplementationName in aURL.Protocol:
            log.debug("is %s in attributes? %s", aURL.Path,
                      hasattr(self, aURL.Path))
            if hasattr(self, aURL.Path):
                return self

    def queryDispatches(self, seqDescripts):
        lDispatcher = []
        for d in seqDescripts:
            lDispatcher.append(self.queryDispatch(d.FeatureURL,
                                                  d.FrameName,
                                                  d.SearchFlags))
        return lDispatcher

    def dispatch(self, aURL, args):
        log.debug("call to dispatch with args %s", args)
        log.debug("aURL.Path = %s", aURL.Path)
        if hasattr(self, aURL.Path):
            getattr(self, aURL.Path)()

    def addStatusListener(self, xControl, aURL):
        pass

    def removeStatusListener(self, xControl, aURL):
        pass
