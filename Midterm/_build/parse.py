import docx,re,json,sys
U="/root/.claude/uploads/4b38b932-a7f5-5fc1-827c-5ae5f103ccb3/"
def chars(p):
    out=[]
    for r in p.runs:
        h=bool(r.font.highlight_color)
        for c in r.text: out.append((c,h))
    return out
def parse(f):
    d=docx.Document(U+f)
    sections=[];cur=None;stem=[];opts=[]
    items=[]  # (section, stem, opts{L:text}, ans)
    pend_stem=[];pend_opts={};pend_hl={}
    def flush():
        nonlocal pend_stem,pend_opts,pend_hl
        if pend_opts:
            ans=[k for k in pend_opts if pend_hl.get(k)]
            items.append(dict(sec=cur,stem=" ".join(pend_stem).strip(),opts=dict(pend_opts),ans=ans))
        pend_stem=[];pend_opts={};pend_hl={}
    for p in d.paragraphs:
        cs=chars(p); t="".join(c for c,_ in cs).replace('\xa0',' ')
        if not t.strip(): continue
        if re.match(r'\s*(PRACTICE|PRATICE)\s+TEST',t,re.I):
            flush(); cur=t.strip(); continue
        if cur is None: continue
        if re.match(r'\s*(Directions|Example)',t): continue
        # find option markers
        ms=list(re.finditer(r'(?:(?<=\s)|^)([A-D])\s?\.\s',t))
        # require sequence starts at A for new item or continues
        if ms and ms[0].group(1)=='B' and not pend_opts:
            # first option 'A.' marker missing: text before B is option A
            cs=[(c,h) for c,h in cs]
            pre=t[:ms[0].start()]
            t="A. "+t.lstrip(); off=3+(len(pre)-len(pre.lstrip()))*0
            lead=len(chars(p)) and 0
            cs=[('A',False),('.',False),(' ',False)]+[(c,h) for c,h in chars(p)]
            t="".join(c for c,_ in cs).replace('\xa0',' ')
            ms=list(re.finditer(r'(?:(?<=\s)|^)([A-D])\s?\.\s',t))
        if ms and (ms[0].group(1)=='A' or (pend_opts and ms[0].group(1) not in pend_opts)):
            pre=t[:ms[0].start()].strip()
            if ms[0].group(1)=='A' and pend_opts: flush()
            if pre: pend_stem.append(pre)
            for i,m in enumerate(ms):
                end=ms[i+1].start() if i+1<len(ms) else len(t)
                seg=t[m.end():end]
                hs=[cs[j][1] for j in range(m.end(),end) if not cs[j][0].isspace()]
                hl=len(hs)>0 and sum(hs)>=len(hs)/2
                # marker itself highlighted?
                if not hl:
                    hs2=[cs[j][1] for j in range(m.start(),m.end()) if not cs[j][0].isspace()]
                    hl=bool(hs2) and all(hs2)
                pend_opts[m.group(1)]=seg.strip(); pend_hl[m.group(1)]=pend_hl.get(m.group(1),False) or hl
        else:
            if pend_opts: flush()
            pend_stem.append(t.strip())
    flush()
    return items
if __name__=="__main__":
    for f in ["04f78d0b-GRAMMAR_MID_L_N_1.docx","93562418-GRAMMAR_MID_L_N_2.docx"]:
        it=parse(f)
        print(f,len(it))
        from collections import Counter
        print(Counter(i['sec'][:70] for i in it))
        bad=[i for i in it if len(i['ans'])!=1 or len(i['opts'])!=4]
        print("bad",len(bad))
        for b in bad[:40]: print(json.dumps(b,ensure_ascii=False))
