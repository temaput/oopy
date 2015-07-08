"""
Word Frequency parser for OO
Copyright © 2015 Artem Putilov

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


def process_word(word):
    punctuation = "!?,.;:)}]"
    starting_brackets = "({["
    remove_punctuation = "".maketrans("", "", punctuation)
    remove_starting_brackets = "".maketrans("", "", starting_brackets)

    if len(word) >= 4:
        word_tuple = (word[0], word[1:-2], word[-2:])
    elif len(word) == 3:
        word_tuple = (word[0], word[1:-1], word[-1])
    else:
        return word.translate(remove_punctuation)

    word_tuple = (
        word_tuple[0].translate(remove_starting_brackets),
        word_tuple[1],
        word_tuple[2].translate(remove_punctuation)
    )
    return "".join(word_tuple)


def freq_report(sourcedoc, targetdoc=None):
    """
    parses writer doc and reports the frequency of words using
    """
    if targetdoc is None:
        targetdoc = sourcedoc
    stat_dict = {}
    from writer import CursorUtilities
    cu = CursorUtilities(sourcedoc)
    for p in cu.iterateParagraphs():
        s = p.String
        for word in s.split():
            word = process_word(word.lower())
            stat_dict[word] = stat_dict.setdefault(word, 0) + 1

    stat_list = [(stat_dict[key], key) for key in stat_dict]
    stat_list.sort(key=lambda t: (-t[0], t[1]))
    from writer import TextUtilities
    tu = TextUtilities(targetdoc)
    for i, word in stat_list:
        tu.appendPara("%s\t%s" % (i, word))


def accumulate_problematic_symbols(encoding, srange, accumulator):
    """
    compose string of unicode symbols in srange, try to encode it
    with encoding, accumulate erroreous symbols in accumulator
    """
    def accumulate_errors(acc):
        def handle_error(err):
            for char in err.object[err.start:err.end]:
                acc.append(ord(char))
            return ("?", err.end)
        return handle_error
    import codecs
    codecs.register_error('accumulate_errors', accumulate_errors(accumulator))
    "".join([chr(i) for i in srange]).encode(
        encoding, errors='accumulate_errors')
    return [r"\u%04x" % i for i in set(srange).difference(set(accumulator))]


def construct_whitelist_search_range(encoding):
    """
    do an opposite to accumulate_problematic_symbols:
        - scan all available symbols in encoding (onebyte)
        - make a sorted set of them
        - compose a search regex string with ranges to exclude all
        available symbols from search and therefore look for porblematic
        only
    """
    from utils import range_creator
    whitelist_range = range_creator(
        sorted(set(ord(char) for char in bytes(range(256)).decode(
            encoding, errors='ignore'))))
    l = []
    for a, b in whitelist_range:
        if a != b:
            l.append(r"\u%04X-\u%04X" % (a, b))
        else:
            l.append(r"\u%04X" % a)

    s = "[^%s]" % "".join(l)
    return s
