#! /usr/bin/env python3

import re


STRUCT_RE = re.compile(r'structure.*: (.*)')
OVERFULL_RE = re.compile(r'Overfull[^(]*\(([0-9.]+)pt')

overfulls = []
struct = None
w = None
for line in open('FémininPluriel_TanguyKervran.log'):
#    print(line)
    if w is not None:
        l = line.split(')', 1)[1]
        overfulls.append((w,struct,l))
        w = None
    else:
        m = STRUCT_RE.match(line)
        if m:
            struct = m.group(1)
            struct
        m = OVERFULL_RE.match(line)
        if m:
            w = float(m.group(1))

overfulls.sort(reverse=True)
for w,s,l in overfulls:
    print("{}\n{}\n{}\n\n".format(w,s,l))
