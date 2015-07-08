import logging
log = logging.getLogger("pyuno")
log.addHandler(logging.NullHandler())


import macrohelper
basic = None
doc = None

if 'XSCRIPTCONTEXT' in globals():
    basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa
    doc = basic.ThisComponent


def set_globals(ctx=None):
    global basic
    global doc
    if ctx is None and 'XSCRIPTCONTEXT' in globals():
        ctx = XSCRIPTCONTEXT  # noqa
    basic = macrohelper.StarBasicGlobals(ctx)
    doc = basic.ThisComponent


from com.sun.star.text.ControlCharacter import (  # noqa
                                                PARAGRAPH_BREAK,
                                                LINE_BREAK,
                                                HARD_HYPHEN,
                                                SOFT_HYPHEN,
                                                HARD_SPACE,
                                                APPEND_PARAGRAPH)


def prepare_for_ventura():
    set_globals()
    ctx = basic.GetDefaultContext()
    from practica import VenturaPrepare
    VenturaPrepare(doc, ctx)()


def convert_index_markers():
    """
    convert bad im entries (index number like <111>) to
    good (index number like {111}
    """
    set_globals()
    ctx = basic.GetDefaultContext()
    from practica import VenturaPrepare
    vp = VenturaPrepare(doc, ctx)
    vp.convert_index_markers()
