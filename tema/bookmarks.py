""" working with bookmarks"""
import logging
log = logging.getLogger(__name__)

import macrohelper
basic = macrohelper.StarBasicGlobals(XSCRIPTCONTEXT)  # noqa

from pythonize import wrapUnoContainer, UnoDateConverter


def addBookmark():
    """ just inserts simple bookmark"""
    doc = basic.ThisComponent
    text = doc.Text
    cursor = text.createTextCursor()
    cursor.gotoStart(False)
    cursor.goRight(4, True)

    bookmark = doc.createInstance("com.sun.star.text.Bookmark")
    bookmark.Name = "Tema test bookmark"
    text.insertTextContent(cursor, bookmark, False)
