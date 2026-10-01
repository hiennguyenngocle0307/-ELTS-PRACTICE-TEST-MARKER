# -*- coding: utf-8 -*-
import json,os,re
HERE=os.path.dirname(os.path.abspath(__file__))
from fixes import *
def load():
    T=json.load(open(os.path.join(HERE,'source_parsed.json'),encoding='utf8'))
    for t in T:
        t['key']={int(k):v for k,v in t['key'].items()}
        t['p1_directions']=P1_DIRECTIONS
        for p in t['p1']:
            if (t['n'],p['n']) in LABEL_FIX: p['label']=LABEL_FIX[(t['n'],p['n'])]
        if t['n'] in SCRIPT2_FIX: t['script2']=SCRIPT2_FIX[t['n']]
        for old,new in SCRIPT1_REPL.get(t['n'],[]):
            assert any(old in l for l in t['script1']),(t['n'],old)
            t['script1']=[l.replace(old,new) for l in t['script1']]
        for q in t['mc']:
            if (t['n'],q['n']) in MC_REORDER:
                order=MC_REORDER[(t['n'],q['n'])]; old=q['opts']; ans=t['key'][q['n']]
                new={L:old[o] for L,o in zip('ABC',order)}
                t['key'][q['n']]=[L for L,o in zip('ABC',order) if o==ans][0]; q['opts']=new
        # cue strings
        t['cues']={}
        for l in t['script1']+t['script2']:
            for m in re.finditer(r'\[Q(\d+): ([^\]]*)\]',l): t['cues'][int(m.group(1))]=m.group(2)
    return T
def sentence_of(t,qn):
    """câu chứa đáp án (đoạn văn bản đến dấu [Qn])."""
    for l in t['script1']+t['script2']:
        m=re.search(r'\[Q%d: [^\]]*\]'%qn,l)
        if m:
            before=l[:m.start()]
            before=re.sub(r'\[Q\d+: [^\]]*\]','',before)
            parts=re.split(r'(?<=[.?!])\s+',before.strip())
            return parts[-1] if parts else before
    return ''
