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
p[style-name='Partie'] => part:fresh
p[style-name='Heading 1'] => chapter:fresh
p[style-name='Subtitle'] => chaptercont:fresh
p[style-name='Heading 2'] => section:fresh
p[style-name='SousTitre2'] => sectioncont:fresh
p[style-name='récit'] => story > p:fresh

r[style-name='footnote reference'] => fr
"""
#p[style-name='footnote text'] => ft

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
        footnotes[num] = note.lstrip('*0123456789 \t')
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
        warnings.warn("unused footnote(s):\n{}".format("\n".join("{}:{}".format(k,v) for k,v in footnotes.items())))
    
    print()
    for filename in h2t.filenames:
        autofile = os.path.join(autodir, filename)
        modiffile = os.path.join(modifdir, filename)
        if os.path.exists(modiffile):
            print()    
            print('\\input {}'.format(modiffile))
            print('%\\input {}'.format(autofile))
            print()    
        else:
            print('\\input {}'.format(autofile))

#\startpart[title={Pluriel\\Mère et enfant},list={Pluriel}]
#\startchapter[title={Etre et ne plus être\\\tfx \it La vie - après la vie - avant la vie},list={Etre et ne plus être}]

#<chapter><a id="_Toc503980580"></a>Etre et ne plus être</chapter>
#<subtitle1>La vie - après la vie - avant la vie</subtitle1>
#<p></p>
#<p>Depuis ...</p>
#<p>Le ...</p>
#<p><em>- Les ...</em></p>

            
class Html2Tex(html.parser.HTMLParser):
    def __init__(self, footnotes, resultdir):
        html.parser.HTMLParser.__init__(self)
        self.footnotes = footnotes
        self.footnote = None
        self.resultdir = resultdir
        self.levels = []
        self.chunks = []
        self.previous = None
        self.ignoredata = 0
        self.filenames = []
        self.numdoc = 0
        self.title = ''
        
    def handle_starttag(self, tag, attrs):
        currentlevel = None if not self.levels else self.levels[-1]
        if (currentlevel in ('part', 'chapter') and tag not in ('a', 'fr')) or\
           (currentlevel!='html' and tag in ('part', 'chapter', 'section')):
            raise Exception("unexpected tag <{}> in <{}>".format(tag, currentlevel))

        self.levels.append(tag)

        if tag == 'html':
            pass
        
        elif tag == 'part':
            if self.previous == 'part':
                self.chunks.pop()
                self.chunks.append('\\\\')
            else:
                self.flush()
                self.chunks.append('\\part{')

        elif tag == 'chapter':
            self.flush()
            self.chunks.append('\\startchapter[list={')

        elif tag == 'chaptercont':
            if self.previous != 'chapter':
                raise Exception("unexpected tag chaptercont after <{}>".format(self.previous))
            self.chunks[-1] = self.chunks[-1][:-3]
            self.chunks.append(r'\\\tfx \it ')

        elif tag == 'section':
            if self.chunks and self.chunks[-1] == '\\\\\n':
                self.chunks.pop()
            self.chunks.append('\\section{')

        elif tag == 'sectioncont':
            if self.previous != 'section':
                raise Exception("unexpected tag sectioncont after <{}>".format(self.previous))
            self.chunks.pop()
            self.chunks.append(r'\\\tfx \it ')

        elif tag == 'story':
            if self.chunks and self.chunks[-1] == '\\\\\n':
                self.chunks.pop()
            self.chunks.append('\\story{\\noindent ')

        elif tag == 'br':
            pass
        
        elif tag == 'p':
            if self.previous == 'p':
                text = ''.join(self.chunks).strip()
                if text.endswith('\\crlf'):
                    while not self.chunks.pop().startswith('\\crlf'):
                        pass
                    self.chunks.append('\n\\blank[4mm]\n')
                else:
                    self.chunks.append('\\crlf\n')
#            self.chunks.append('\\indent ')
            pass
        
        elif tag == 'ft':
            pass

        elif tag == 'ul':
            self.chunks.append('\\startitemize[2,packed,paragraph,intro]\n')

        elif tag == 'li':
            self.chunks.append('\\item ')
            
        elif tag == 'em':
            self.chunks.append('{\\em ')
            
        elif tag == 'strong':
            self.chunks.append('{\\bf ')

        elif tag == 'sup':
            self.chunks.append('{\\high ')

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

        
        else:
            raise Exception("unexpected tag <{}>".format(tag))

    def handle_data(self, data):
        assert self.ignoredata >= 0
        if not self.ignoredata:
            if data.startswith(' ') and self.chunks[-1].startswith('{\\'):
                self.chunks[-1] = ' ' + self.chunks[-1]
            self.chunks.append(data)

    def handle_endtag(self, tag):
        self.previous = self.levels.pop()
        
        if tag == 'html':
            self.flush()

        elif tag == 'part':
            if not self.title:
                self.title = ''.join(self.chunks[1:])
            self.chunks.append('}\n')

        elif tag == 'chapter':
            self.title = ''.join(self.chunks[1:])
            #title={Etre et ne plus être\\\tfx \it La vie - après la vie - avant la vie},list={Etre et ne plus être}]
            self.chunks.append('}},title={{{0}}}]\n'.format(self.title))
            
        elif tag == 'chaptercont':
            self.chunks.append('}]\n')

        elif tag in ('section', 'story'):
            if self.chunks[-1] == '\\\\\n':
                self.chunks.pop()
            self.chunks.append('}\n')
            
        elif tag == 'sectioncont':
            self.chunks.append('}\n')

        elif tag == 'br':
            pass

        elif tag == 'p' or tag == 'ft':
##            if self.chunks[-1] == '\\indent ':
##                self.chunks.pop()
##                text = ''.join(self.chunks).strip()
##                if text.endswith('\\\\'):#self.chunks and self.chunks[-1] == '\\\\\n':
##                    while self.chunks:
##                        chunk = self.chunks.pop()
##                        if chunk == '\\\\\n':
##                            break
##                    self.chunks.append('\n\n')
##            else:
#                self.chunks.append('\\\\\n')
            pass

        elif tag == 'ul':
            self.chunks.append('\\stopitemize\n')

        elif tag == 'li':
            self.chunks.append('\n')

        elif tag in ('em', 'strong', 'sup'):
            if self.chunks[-1].startswith('{\\'):
                self.chunks.pop()
            else:
                self.chunks.append('}')

        elif tag == 'fr':
            self.ignoredata -= 1
            if not self.footnote:
                warnings.warn('empty footnote ...{}*'.format(self.chunks[-5:-1]))
            else:
                self.chunks.append('\\footnote{{{}}}'.format(self.footnote))
            self.footnote = None

        elif tag == 'a':
            self.ignoredata -= 1
                    
    def flush(self):
        self.numdoc += 1
        print(self.numdoc, self.title)
        name = '_'.join(self.title.translate(str.maketrans('', '', ',./-"\\(){}?')).split())
        filename = 'doc{:02}-{}.tex'.format(self.numdoc, name[:15].strip('_'))
        doc = open(os.path.join(self.resultdir, filename), 'w', encoding='utf-8')
        text = ''.join(self.chunks).replace('%', '\\%').replace('$', '\\$').rstrip("\\\n")
        doc.write(text)
        doc.close()
        self.filenames.append(filename)
        self.chunks = []
        self.title = ''


open('check.html','w').write(htmldoc)
html2tex(htmldoc, footnotes, 'chapitres')
