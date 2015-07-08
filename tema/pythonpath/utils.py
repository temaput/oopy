""" my version of summerfield utils,
some usefull funcs for instrumental"""

import logging
log = logging.getLogger("pyuno.utils")

from itertools import groupby

#
# tools
#


def range_creator(iterable):
    """
    1,2,3 5,7,8,9,10 --> (1,3), (5, 5), (7, 10)
    """
    for _, couplist in groupby(enumerate(iterable),
                               lambda coup: coup[1] - coup[0]):
        couplist = list(couplist)
        yield couplist[0][1], couplist[-1][1]


#
# class decorators
#


def delegate(methodnames, wrapper=None, attrname=None):
    """ decorates class with methods in methodnames
        methods are references to the same methods in
        attrname object (if given), or in parent class
        if wrapper (func reference) is given
        than the result of these methods will be passed to it first

        using this class decorator makes the effect of decorating each
        class method i.e.

        @delegate(('add', 'sub'), wrapper, 'obj')
        class...:

        is the same as to say

        class ...:
        @wrapper
        def add(self):
            return obj.add()

        @wrapper
        def sub(self):
            return obj.sub()
    """
    if wrapper is None and attrname is None:
        return lambda cls: cls
    result_wrapper = wrapper or (lambda result: result)
    def func_decorator(cls, mname):
        def func_wrapper(self, *args, **kwargs):
            return result_wrapper(
                getattr(
                getattr(self, attrname) if attrname else super(cls, self),
                mname)(*args, **kwargs))
        return func_wrapper
    def class_decorator(cls):
        for mname in methodnames:
                setattr(cls, mname, func_decorator(cls, mname))
        return cls
    return class_decorator


class Bunch:
    def __init__(self, **kwargs):
        self.__dict__.update(kwargs)


#
# sys utils
#

def removeany(path, ignore_errors=False):
    from os.path import isdir, exists
    from shutil import rmtree
    from os import remove
    if not exists(path):
        return
    try:
        if isdir(path):
            rmtree(path)
        else:
            remove(path)
    except OSError as e:
        if not ignore_errors:
            raise
        else:
            return e.strerror
