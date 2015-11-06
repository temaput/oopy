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
    from practica import VenturaPrepare
    VenturaPrepare(basic)()


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


def freq_report():
    """
    Create frequency report
    """
    from freq import freq_report
    sourcedoc = basic.macro_create_doc("writer")
    freq_report(basic.ThisComponent, sourcedoc)
    basic.MsgBox("Done!")


def print_index_from_doc():
    """
    Prints index from editor doc
    """
    import writer
    iu = writer.IndexUtilities2(doc)
    target = basic.macro_create_doc("writer")
    iu.printIndex(target)


def print_index_from_layout():
    """
    Print index from exported index from layout
    """
    from indexmaker import IndexMaker
    target = basic.macro_create_doc("writer")
    IndexMaker(doc, target)()


def expand_table():

    from practica import expand_table, VenturaPrepare
    vp = VenturaPrepare(basic)
    vp.symbol_substitute()
    expand_table(basic, ",")


def reorder_bibliography():

    from practica import BibliographyReorder
    br = BibliographyReorder(doc)
    br.do_reorder()

__all__ = (prepare_for_ventura,
           convert_index_markers,
           freq_report,
           )
