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

from com.sun.star.frame import (XDispatchProvider,
                                XDispatch)


class HelloWorldProtocolHandler(unohelper.Base, XDispatchProvider, XDispatch):
    """
    Implements ProtocolHandler
    component as in https://wiki.openoffice.org/wiki/Documentation/DevGuide/
    WritingUNO/Implementation
    """
    implementationName = "org.openoffice.comp.pyuno.demo.HelloWorld"

    def __init__(self, ctx):
        # store the component context for later use
        self.ctx = ctx

    def dothejob(self):
        # note: args[0] == "HelloWorld", see below config settings

        # retrieve the desktop object
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx)

        # get current document model
        model = desktop.getCurrentComponent()

        # access the document's text property
        text = model.Text

        # create a cursor
        cursor = text.createTextCursor()

        # insert the text into the document
        text.insertString(cursor, "Hello World", 0)

    def dothejob2(self):
        # note: args[0] == "HelloWorld", see below config settings

        # retrieve the desktop object
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx)

        # get current document model
        model = desktop.getCurrentComponent()

        # access the document's text property
        text = model.Text

        # create a cursor
        cursor = text.createTextCursor()

        # insert the text into the document
        text.insertString(cursor, "Hello World2", 0)

    def queryDispatch(self, aURL, sTargetFrameName, iSearchFlags):
        print("in queryDispatch")
        log.debug("looking for %s in %s", self.implementationName,
                  aURL.Protocol)
        if self.implementationName in aURL.Protocol:
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

# pythonloader looks for a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    HelloWorldProtocolHandler,                        # UNO object class
    HelloWorldProtocolHandler.implementationName,  # implementation name
    ("com.sun.star.frame.ProtocolHandler",),)  # list of implemented services
# (the only service)
