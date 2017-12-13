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


# ====== docx to html ======
style_map = """
p[style-name='Heading 1'] => chapter:fresh
p[style-name='Heading 2'] => section:fresh
p[style-name='Récit'] => story > p:fresh

p[style-name='footnote text'] => ft
r[style-name='footnote reference'] => fr
"""

with open(sys.argv[1], 'rb') as docx_file:
    result = mammoth.convert_to_html(docx_file, style_map=style_map, ignore_empty_paragraphs=False)
for m in result.messages:
    print(m)

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
htmldoc = '<html>{}</html>'.format(doc)



# ====== html to tex ======
def html2tex(htmldoc, footnotes, resultdir):
    autodir = os.path.join(resultdir, 'auto')
    modifdir = os.path.join(resultdir, 'modif')
    for autofile in glob.glob(os.path.join(autodir, '*.tex')):
        os.remove(autofile)
        
    h2t = Html2Tex(footnotes, autodir)
    h2t.feed(htmldoc)
    if footnotes:
        raise Exception("unused footnote(s):\n{}".format("\n".join(footnotes.values())))
    
    print()
    for filename in h2t.filenames:
        autofile = os.path.join(autodir, filename)
        modiffile = os.path.join(modifdir, filename)
        if os.path.exists(modiffile):
            print('\\input {}'.format(modiffile))
            print('%\\input {}'.format(autofile))
        else:
            print('\\input {}'.format(autofile))
        print()    

class Html2Tex(html.parser.HTMLParser):
    def __init__(self, footnotes, resultdir):
        html.parser.HTMLParser.__init__(self)
        self.footnotes = footnotes
        self.footnote = None
        self.resultdir = resultdir
        self.levels = []
        self.chunks = []
        self.ignoredata = 0
        self.filenames = []
        self.doc = None
        self.numdoc = 0
        
    def handle_starttag(self, tag, attrs):
        currentlevel = None if not self.levels else self.levels[-1]
        if (currentlevel in ('part', 'chapter', 'section') and tag != 'a') or\
           (currentlevel!='html' and tag in ('part', 'chapter', 'section')):
            raise Exception("unexpected tag <{}> in <{}>".format(tag, currentlevel))

        self.levels.append(tag)

        if tag == 'html':
            pass
        
        elif tag in ('part', 'chapter', 'section', 'story'):
            if tag in ('part', 'chapter'):
                self.flush()
            self.chunks.append('\\{}{{'.format(tag))

        elif tag == 'p':
            self.chunks.append('\\indent ')
            
        elif tag == 'em':
            self.chunks.append('{\\em ')
            
        elif tag == 'strong':
            self.chunks.append('{\\bf ')

        elif tag == 'fr':
            self.ignoredata += 1
            
        elif tag == 'a':
            self.ignoredata += 1
            if 'fr' in self.levels:
                notenum = dict(attrs)['id'][len('footnote-ref-'):]
                try:
                    self.footnote = self.footnotes.pop(notenum)
                except:
                    raise Exception("missing note number {} in footnotes".format(notenum))

        elif tag == 'sup':
            pass
        
        else:
            raise Exception("unexpected tag <{}>".format(tag))

    def handle_data(self, data):
        if not self.ignoredata:
            self.chunks.append(data)

    def handle_endtag(self, tag):
        self.levels.pop()
        
        if tag == 'html':
            self.flush()
            if self.doc:
                self.doc.close()
                self.doc = None

        elif tag in ('part', 'chapter', 'section', 'story'):
            if tag in ('part', 'chapter'):
                title = ''.join(self.chunks[1:]).strip().replace(' ','_')
                if not title:
                    self.chunks = []
                else:
                    self.numdoc += 1
                    filename = 'doc{:02}-{}-{}.tex'.format(self.numdoc, tag, title[:15])
                    if self.doc:
                        self.doc.close()
                    self.doc = open(os.path.join(self.resultdir, filename), 'w', encoding='utf-8')
                    print(''.join(self.chunks[1]))
                    self.filenames.append(filename)
                    self.chunks.append('}\n')
            else:
                self.chunks.append('}\n')

        elif tag == 'p':
            self.chunks.append('\\\\\n')


        elif tag in ('em', 'strong'):
            self.chunks.append('}')

        elif tag == 'fr':
            self.ignoredata -= 1
            if not self.footnote:
                raise Exception('empty footnote')
            self.chunks.append('\\footnote{{{}}}'.format(self.footnote))
            self.footnote = None

        elif tag == 'a':
            self.ignoredata -= 1
                    
    def flush(self):
        text = ''.join(self.chunks)
        if self.doc:
            self.doc.write(text)
        elif text.strip():
#            raise Exception("no open doc to flush {!r}".format(text))
            warnings.warn("no open doc to flush {!r}".format(text))
        self.chunks = []


open('check.html','w').write(htmldoc)
html2tex(htmldoc, footnotes, 'chapitres')
