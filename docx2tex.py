#! /usr/bin/env python3
# coding: utf-8

#    if p.style.name == 'Heading 1':
#        title = p.text.strip()
#        if not title:
#            continue
#        chapnum += 1
#        chaptitle = title.replace(' ', '_')
#        print(chapnum, chaptitle)
#        chap = open('result/{:02}-{}.txt'.format(chapnum, chaptitle), 'w')
#        chap.write('\chapter{{{}}}\n'.format(title))

import sys
import glob
import shutil
import os
import re
import xml.etree.ElementTree as ET
import html.parser
import warnings

import mammoth


style_map = """
p[style-name='Récit'] => story:fresh
p[style-name='footnote text'] => ft
r[style-name='footnote reference'] => fr
"""
#p[style-name='endnote text'] => et
#r[style-name='endnote reference'] => er

with open(sys.argv[1], 'rb') as docx_file:
    result = mammoth.convert_to_html(docx_file, style_map=style_map)
for m in result.messages:
    print(m)

#open('result/livre.html','w').write("""<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN"
#  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
#<html xmlns="http://www.w3.org/1999/xhtml">
#  <head>
#    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
#    <title>XXX</title>
#  </head>
#  <body>
#{}
#  </body>
#</html>""".format(result.value))

doc,sep,footnotes = result.value.rpartition('<ol>')
if not footnotes.endswith('</ol>'):
    warnings.warn("No footnotes")
    doc = result.value
    footnotes = {}
else:
    #<fr><sup><a href="#footnote-84" id="footnote-ref-84">[84]</a></sup>*</fr>
    #<li id="footnote-84"><ft><fr>*</fr> Peter Chamberlain. 1560 - 1631</ft><p> <a href="#footnote-ref-84">↑</a></p></li>
    tree = ET.fromstring(sep+footnotes)
    footnotes = {}
    for li in tree.findall('.//li'):
        num = li.attrib['id'][len('footnote-'):]
        note = ''.join(li.itertext()).strip('* ↑\n')
        footnotes[num] = note


class Html2Tex(html.parser.HTMLParser):
    def __init__(self, htmldoc, resultdir):
        html.parser.HTMLParser.__init__(self)
        self.numdoc = 0
        self.doc = None
        self.opendoc()
        self.levels = []
        self.texts = []

        for autofile in glob.glob(os.path.join(resultdir, 'auto', '*.tex')):
            os.remove(autofile)
        self.feed(htmldoc)
        if self.doc:
            self.doc.close()
        for autofile in glob.glob(os.path.join(resultdir, 'auto', '*.tex')):
            basename = os.path.basename(autofile)
            modiffile = os.path.join(resultdir, 'modif', basename)
            reffile = os.path.join(os.pardir, 'auto', basename)
            if not os.path.exists(modiffile):
                os.symlink(reffile, modiffile)
            
    def opendoc(self):
        if self.doc:
            self.doc.close()
        self.doc = open('chapitres/auto/chap{:02}.tex'.format(self.numdoc), 'w', encoding='utf-8')
        self.numdoc += 1

        
    def writetext(self, before=None, after=None, text=None):
        if text is None:
            text = ''.join(self.texts).replace('\n', ' ')
        if text.strip():
            self.doc.write("{}{}{}".format(before or '', text, after or ''))
            self.texts = []
        
    def handle_starttag(self, tag, attrs):
        self.writetext()
        
        self.levels.append(tag)
        if tag == 'h1':
            assert(len(self.levels) == 1)
        elif tag == 'h2':
            assert(len(self.levels) == 1)
        elif tag == 'p':
            assert(len(self.levels) == 1)
#        elif tag == 'ul':
#            self.writetext('', '', '\n\n')
            assert(len(self.levels) == 1)
        elif tag == 'fr':
            assert(len(self.levels) == 2)
            self.notenum = None
        elif tag == 'a':
            if len(self.levels) >= 3 and self.levels[-3] == 'fr':
                self.notenum = dict(attrs)['id'][len('footnote-ref-'):]

    def handle_endtag(self, tag):
        if tag == 'h1':
            self.opendoc()
            self.writetext('\\chapter{', '}\n')
        
        elif tag == 'h2':
            self.writetext('\\section{', '}\n')
        
        elif tag == 'p':
            self.writetext('', '\n\n')
        
        elif tag == 'story':
            self.writetext('\\story{', '}\n')

        elif tag == 'fr':
            if self.notenum is not None:
                footnote = footnotes[self.notenum]
                self.writetext('\\footnote{', '}', footnote)

        elif tag == 'a':
            pass

        elif tag == 'ft':
            pass

        elif tag == 'sup':
            pass

        elif tag == 'em':
            self.writetext('\\emph{', '}')

        elif tag == 'strong':
            self.writetext('{\\bf ', '}')

        elif tag == 'li':
            self.writetext(' * ', '\n')

        elif tag == 'ul':
            self.writetext('', '', '\n\n')

        elif tag == 'br':
            self.writetext('', '', '\n')

        else:
            raise Exception("tag {} ignored".format(tag))

        self.levels.pop()

    def handle_data(self, data):
        self.texts.append(data)

open('check.html','w').write(doc)
Html2Tex(doc, 'chapitres')
