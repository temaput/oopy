""" Pythonic wrappers that translate JAVA-like interfaces to more pythonic
    analogs """

import uno
from com.sun.star.lang import IndexOutOfBoundsException


# ---------------------------------------------------------------------------
#           uno containers conversions and wrappers
# ---------------------------------------------------------------------------
import sys

def wrapUnoContainer(UnoContainter):
    """ magic: picks the right convertor for any type of uno Containers"""
    for t in UnoContainter.Types:
        branch, interface = t.typeName.split('.')[3:5]
        if branch == 'container':
            this = sys.modules[__name__]
            return this.__dict__[interface](UnoContainter)
    raise (TypeError, "Bad uno container type")


class XEnumerationAccess:
    """ translates com.sun.star.container.XEnumerationAccess interface
    object to iterator"""

    def __init__(self, Xobject):
        self._object = Xobject

    def __iter__(self):
        xEnum = self._object.createEnumeration()
        while(xEnum.hasMoreElements()):
            yield xEnum.nextElement()


class XIndexAccess:
    """translates com.sun.start.container.XIndexAccess to list"""

    def __init__(self, Xobject):
        self._object = Xobject

    def __getitem__(self, index):
        return self._object.getByIndex(index)

    def __setitem__(self, index, item):
        try:
            self._object.replaceByIndex(index, item)

        except IndexOutOfBoundsException as e:
            raise (IndexError, e)


class XNameAccess(dict):
    """translates com.sun.start.container.XIndexAccess to dict"""

    def __init__(self, Xobject):
        self._object = Xobject

    def __getitem__(self, key):
        return self._object.getByName(key)

    def __setitem__(self, key, value):
        self._object.setByName(key, value)

    def keys(self):
        return self._object.getElementNames()

    def __contains__(self, key):
        return self._object.hasByName(key)

    def __iter__(self):
        return (key for key in self.keys())

    def values(self):
        return [self[key] for key in self.keys()]

    def items(self):
        return[(key, self[key]) for key in self]


# -------------------------------------------------
#           date / time conversions
# -------------------------------------------------
from datetime import datetime
from com.sun.star.util import Date as unoDate
from com.sun.star.util import DateTime as unoDateTime
from tema.utils import delegate

def _wrapspecials(result):
    if type(result) == datetime:
        return UnoDateConverter.fromDateTime(result)
    else:
        return result

@delegate(('__add__', '__radd__', '__sub__', '__rsub__'), _wrapspecials)
class UnoDateConverter(datetime):
    """ Makes easy conversion between python datetime.datetime and
    OO com.sun.star.util Date and DateTime 

    it wraps python native datetime:
    >>> UnoDateConverter.today().day == datetime.today().day
    True

    it can be used in +- expressions aswell
    >>> from datetime import timedelta
    >>> UnoDateConverter(2013, 10, 16) + timedelta(days=-1)
    UnoDateConverter(2013, 10, 15, 0, 0)
    >>> UnoDateConverter(2013, 10, 16) - datetime(2013, 10, 15)
    datetime.timedelta(1)
    >>> datetime(2013, 10, 16, 15, 55) - UnoDateConverter(2013, 10, 16, 15, 50)
    datetime.timedelta(0, 300)

    it returns UnoDate and UnoDateTime
    >>> UnoDateConverter(2013, 10, 16).getUnoDate()
    (com.sun.star.util.Date){ Day = (unsigned short)0x10, Month = (unsigned short)0xa, Year = (short)0x7dd }
    >>> UnoDateConverter(2013, 10, 16).getUnoDateTime()
    (com.sun.star.util.DateTime){ NanoSeconds = (unsigned long)0x0, Seconds = (unsigned short)0x0, Minutes = (unsigned short)0x0, Hours = (unsigned short)0x0, Day = (unsigned short)0x10, Month = (unsigned short)0xa, Year = (short)0x7dd, IsUTC = (boolean)false }
    
    it can be created from existing datetime instance:
    >>> dt = datetime.today()
    >>> UnoDateConverter.fromDateTime(datetime(2013, 10, 16)).day
    16
    
    it can be created from existing unoDate or unoDateTime aswell:
    >>> UnoDateConverter.fromUnoDate(unoDate(16, 10, 2013))
    UnoDateConverter(2013, 10, 16, 0, 0)
    >>> UnoDateConverter.fromUnoDate(unoDateTime(0, 0, 1, 1, 16, 10, 2013, False))
    UnoDateConverter(2013, 10, 16, 1, 1)

    """

    members = (
                ('year', 'Year'),
                ('month', 'Month'),
                ('day', 'Day'),
                ('hour', 'Hours'),
                ('minute', 'Minutes'),
                ('second', 'Seconds'),
                ('microsecond', 'NanoSeconds', 1000)
                )

    def getUnoDate(self):
        return unoDate(self.day, self.month, self.year)

    def getUnoDateTime(self):
        unodict = {}
        for member in self.members:
            value = getattr(self, member[0])
            if len(member) > 2:  # we have to apply convertions ratio
                value *= member[2]
            unodict[member[1]] = value
        unodict['IsUTC'] = False  # not implemented yet
        return unoDateTime(**unodict)

    @classmethod
    def fromUnoDate(cls, unoDate):
        dtdict = {}
        for member in cls.members:
            if hasattr(unoDate, member[1]):
                value = getattr(unoDate, member[1])
                if len(member) > 2:  # make additional conversion
                    value //= member[2]
            else:
                value = 0
            dtdict[member[0]] = value
        return cls(**dtdict)

    @classmethod
    def fromDateTime(cls, dt):
        dtdict = {}
        for member in cls.members:
            if hasattr(dt, member[0]):
                value = getattr(dt, member[0])
            else:
                value = 0
            dtdict[member[0]] = value
        return cls(**dtdict)


if __name__ == '__main__':
    import doctest
    doctest.testmod()
