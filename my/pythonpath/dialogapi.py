""" api for creating and handling dialogs progrmmaticaly 
    http://wiki.openoffice.org/wiki/Documentation/DevGuide/GUI/Dialogs_and_Controls
"""

import logging
log = logging.getLogger(__name__)
from tema.SortedDict import SortedDict
from uno import Enum

getEnum = lambda branch, name: Enum("{}.{}".format(branch, name))
makeID = lambda name: "{}{}".format(name, id(name))

class unoDialog(object):
    """ kinda base class for making dialogs programmaticaly
        supposed to be inherited, base props given in setProperties,
        wich should return them as dict.
        Construction takes place in construct method. 
        Base class provides neccesary methods for creating control elements,
        handling events. Child class defins specific dialog. Instance 
        supplies constructor with context, parent window and service mgr and
        starts and disposes dialog, returning values (1 or 0) and saving all
        user input
    """


    def __init__(self, ctx, smgr, parent, **kwargs):
        self.toolkit = parent.getToolkit()
        self.dialog = dialog = smgr.createInstanceWithContext(
                "com.sun.star.awt.UnoControlDialog", ctx)
        self.dialogModel = dialogModel = smgr.createInstanceWithContext(
                "com.sun.star.awt.UnoControlDialogModel", ctx)
        dialog.setModel(dialogModel)
        properties = SortedDict(self.setProperties(**kwargs))
        if properties:
            self.dialogModel.setPropertyValues(tuple(properties.keys()),
                    tuple(properties.values()))
        print ("MOTHERFUCK")
        log.debug("dialog is %s", dialog)
        self.construct(**kwargs)

    def setProperties(self, **kwargs): 
        return dict()
    def construct(self, **kwargs): raise NotImplementedError

    def __insertControl(self, controlType, properties): 
        instance = self.dialogModel.createInstance(
                "com.sun.star.awt.{}".format(controlType))
        log.debug("Instance is %s", instance)
        properties = SortedDict(properties)
        properties['Name'] = name = makeID(controlType)
        instance.setPropertyValues(tuple(properties.keys()),
                    tuple(properties.values()))
        self.dialogModel.insertByName(name, instance)
        return self.dialog.getControl(name) 

    def insertLabel(self, message = "", **properties):
        control = self.__insertControl("FixedText", properties)
        control.setText(message)
        return control

    def insertButton(self, label = "", btnType = "STANDART", eventHandler = None, 
            **properties):
        properties['Label'] = label
        properties['PushButtonType'] = getEnum("com.sun.star.awr.PushButtonType",
                                btnType)
        control = self.__insertControl("PushButton", properties)
        return control

    def show(self):
        self.dialog.createPeer(self.toolkit, self.parent)
        self.dialog.execute()

    #def __del__(self):
    #    self.dialog.dispose()

