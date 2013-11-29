"""
Basic class implements global Constants and Functions in StarBasic
Makes writing macroses with python more comfortable
Takes XSCRIPTCONTEXT or ComponentContext to generate all other objects, so it can  be
used inside macroses and from remote procs

"""

import logging
log = logging.getLogger(__name__)
import uno

import dialogapi


class InputBox(dialogapi.unoDialog):

    def setProperties(self, **kwargs):
        properties = dict(
                Name = 'InputBox'
                )
        properties.update(kwargs)
        return properties

    def construct(self, **kwargs):
        self.labelControl = self.insertLabel(kwargs.get('message', ''),
                PositionX = 5, PositionY = 5, Width = 100, Height = 8)
        self.btnOK = self.insertButton('OK', 'OK', PositionX=100, PositionY=100,
                Height=14, Width=40, DefaultButton=True)
        
class StarBasicGlobals:
    def __init__(self, context):
        self._givenctx = context
        self._preload()

    def _preload(self):
        c = self._givenctx
        if hasattr(c, 'getDesktop'):  # context is XSCRIPTCONTEXT
            self.StarDesktop = c.getDesktop()
            self.GetDefaultContext = c.getComponentContext
            self.ThisComponent = c.getDocument()
        elif hasattr(c, 'ServiceManager'):  # context is StarOffice.ComponentContext
            self.StarDesktop = c.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", c)
            self.GetDefaultContext = lambda: c
            self.ThisComponent = self.StarDesktop.getCurrentComponent()
        assert self.StarDesktop, "bad context, init failed"

        self._ctx = self.GetDefaultContext()
        self.GetProcessServiceManager = self._ctx.getServiceManager
        self._smgr = self.GetProcessServiceManager()
    
    def CreateUnoService(self, uri, *args):
        if len(args):
            return self._smgr.createInstanceWithArgumentsAndContext(uri,
                    args, self._ctx)
        return self._smgr.createInstanceWithContext(uri, self._ctx)
    def CreateUnoStruct(self, uri, *args):
        return uno.createUnoStruct(uri, *args)

    def CreateUnoDialog(self, uri):
        dp = self._smgr.createInstanceWithContext(
                "com.sun.star.DialogProvider", self._ctx)
        return dp.createDialog(
                "vnd.sun.star.script:{uri}?location=user".format(uri=uri))

    def InputBox(self, message, title="Input"):
        """ Create simple imput dialog programaticaly"""
        parent = self.StarDesktop.getCurrentFrame().getContainerWindow()
        inp = InputBox(self._ctx, self._smgr, parent, Title=title)
        inp.show()
        
    def MsgBox(self, message, title="", message_type="infobox", buttons=1):
        """ Show message in message box. """
        parent = self.StarDesktop.getCurrentFrame().getContainerWindow()
        toolkit = parent.getToolkit()
        older_imple = self._check_method_parameter(
                "com.sun.star.awt.XMessageBoxFactory", "createMessageBox",
                1, "com.sun.star.awt.Rectangle")
        if older_imple:
            from com.sun.star.awt import Rectangle
            msgbox = toolkit.createMessageBox(
                parent, Rectangle(), message_type, buttons, title, message)
        else:
            message_type = uno.getConstantByName("com.sun.star.awt.MessageBoxType." + {
                "messbox": "MESSAGEBOX", "infobox": "INFOBOX",
                "warningbox": "WARNINGBOX", "errorbox": "ERRORBOX",
                "querybox": "QUERYBOX"}[message_type])
            msgbox = toolkit.createMessageBox(
                parent, message_type, buttons, title, message)
        n = msgbox.execute()
        msgbox.dispose()
        return n

    def _check_method_parameter(self, interface_name, method_name, param_index, param_type):
        """ Check the method has specific type parameter at the specific position. """
        cr = self.CreateUnoService("com.sun.star.reflection.CoreReflection")
        try:
            idl = cr.forName(interface_name)
            m = idl.getMethod(method_name)
            if m:
                info = m.getParameterInfos()[param_index]
                return info.aType.getName() == param_type
        except:
            pass
        return False

