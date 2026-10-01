import re,json,os
HERE=os.path.dirname(os.path.abspath(__file__))
def parse(path=os.path.join(HERE,'source_youpass_dump.txt')):
    L=[l.rstrip('\n') for l in open(path,encoding='utf8')]
    tests=[];i=0
    starts=[k for k,l in enumerate(L) if re.match(r'^TEST \d+:',l)]
    endk=[k for k,l in enumerate(L) if l.startswith('Suggested 3-Step')][0]
    for si,s in enumerate(starts):
        e=starts[si+1] if si+1<len(starts) else endk
        B=L[s:e]; t={}
        m=re.match(r'^TEST (\d+): (.*)$',B[0]); t['n']=int(m.group(1)); t['title']=m.group(2)
        t['anchor']=B[1].replace('Source anchor: ',''); t['url']=B[2].strip()
        def idx(pred,start=0):
            for k in range(start,len(B)):
                if pred(B[k]): return k
        p1=idx(lambda l:l.startswith('PART 1')); p2=idx(lambda l:l.startswith('PART 2')); ak=idx(lambda l:l.startswith('ANSWER KEY')); tr=idx(lambda l:l.startswith('FULL TRANSCRIPT'))
        t['p1_directions']=B[p1+1]; t['p1_title']=B[p1+2]
        t['p1']=[]
        for l in B[p1+3:p2]:
            m=re.match(r'^(\d+) \| (.*?) \| _+',l)
            if m: t['p1'].append(dict(n=int(m.group(1)),label=m.group(2)))
        t['p2_title']=B[p2+1]
        mc=[];cur=None;match_opts=[];match_items=[];mode=None
        for l in B[p2+2:ak]:
            if l.startswith('Questions 11-15'): mode='mc'; t['mc_directions']=l; continue
            if l.startswith('Questions 16-20'): mode='match'; t['match_directions']=l; continue
            if mode=='mc':
                m=re.match(r'^(\d+)\. (.*)$',l)
                if m: cur=dict(n=int(m.group(1)),q=m.group(2),opts={}); mc.append(cur); continue
                m=re.match(r'^\s+([ABC])\. (.*)$',l)
                if m: cur['opts'][m.group(1)]=m.group(2)
            elif mode=='match':
                m=re.match(r'^([A-E]) - (.*)$',l)
                if m: match_opts.append((m.group(1),m.group(2))); continue
                m=re.match(r'^(\d+)\. (.*?): _+',l)
                if m: match_items.append(dict(n=int(m.group(1)),label=m.group(2)))
        t['mc']=mc; t['match_opts']=match_opts; t['match_items']=match_items
        key={}
        for l in B[ak+1:tr]:
            m=re.match(r'^(\d+) \| (.*?) \| (\d+) \| (.*)$',l)
            if m: key[int(m.group(1))]=m.group(2); key[int(m.group(3))]=m.group(4)
        t['key']=key
        s1=idx(lambda l:l=='Part 1 Script'); s2=idx(lambda l:l=='Part 2 Script'); tf=idx(lambda l:l.startswith('Teacher Focus'))
        t['script1']=[l for l in B[s1+1:s2] if l.strip()]
        t['script2']=[l for l in B[s2+1:tf] if l.strip()]
        tests.append(t)
    return tests
if __name__=="__main__":
    T=parse(); json.dump(T,open(os.path.join(HERE,'source_parsed.json'),'w'),ensure_ascii=False,indent=1)
    for t in T:
        assert len(t['p1'])==10 and len(t['mc'])==5 and len(t['match_items'])==5 and len(t['key'])==20 and len(t['match_opts'])==5,(t['n'],len(t['p1']),len(t['mc']),len(t['match_items']),len(t['key']))
        assert all(len(q['opts'])==3 for q in t['mc'])
        # cue consistency
        cues={}
        for l in t['script1']+t['script2']:
            for m in re.finditer(r'\[Q(\d+): ([^\]]*)\]',l): cues[int(m.group(1))]=m.group(2)
        assert set(cues)==set(range(1,21)),(t['n'],sorted(set(range(1,21))-set(cues)))
        bad=[q for q in cues if cues[q]!=t['key'][q]]
        print(t['n'],t['title'],'cue/key mismatch:',bad)
