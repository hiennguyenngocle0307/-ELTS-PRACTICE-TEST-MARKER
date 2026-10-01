import json,re,os,sol_m1,sol_m2
CFG={
 'm1':dict(file='GRAMMAR_MID_L_N_1',
   modules=[("Module 1","Verb Tenses","Chapter 1",range(1,41)),
            ("Module 2","Singular and Plural","Chapter 5",range(41,81)),
            ("Module 3","Adjective Clauses","Chapter 6",range(81,121)),
            ("Module 4","Noun Clauses","Chapter 7",range(121,161))],
   exclude={102,119,157}, S=sol_m1.S),
 'm2':dict(file='GRAMMAR_MID_L_N_2',
   modules=[("Module 1","Modals and Similar Expressions","Chapter 2",range(1,41)),
            ("Module 2","The Passive","Chapter 3",range(41,81)),
            ("Module 3","Gerunds and Infinitives","Chapter 4",range(82,122)),
            ("Module 4","Connecting Ideas (adverb clauses & transitions)","Chapter 8",range(122,162))],
   exclude={2,89,123,124,133,139}, S=sol_m2.S),
}
OPT_FIX={('m1',43):{'C':'Is there any proofs'},('m1',44):{'D':'many luggage'},('m1',73):{'D':'is a lot of problem'},
         ('m2',97):{'A':'to let me to borrow','D':'let me to borrow'},('m2',92):{'D':'watching to land it'},
         ('m2',98):{'C':'being won'},('m2',120):{'A':'have it painted'},
         ('m2',91):{'C':'try','D':'for trying'},('m2',84):{'B':'involving','C':'having involved'},('m2',64):{'C':'had been discovered'}}
STEM_FIX={('m2',96):'Jack made me ________ him next week.',('m2',98):"I'll never forget ________ that race. What a thrill!",
          ('m2',10):'“Since we have to be there in a hurry, we ________ take a taxi.”\n“I agree.”'}
def norm(t):
    t=t.replace('’',"'").replace('‘',"'").replace('“','"').replace('”','"')
    t=re.sub(r'[ \t]+',' ',t)
    t=re.sub(r'\s+([?!.,;:])',r'\1',t)
    t=re.sub(r'(?<=[a-z])\.(?=[A-Z])',r'. ',t)
    t=re.sub(r'_{2,}|(?<!\w)_(?!\w)','________',t)
    t=re.sub(r'(________)(?=[A-Za-z])',r'\1 ',t); t=re.sub(r'(?<=[A-Za-z])(________)',r' \1',t)
    lines=[]
    for ln in t.strip().split('\n'):
        ln=ln.strip()
        if ln.count('"')%2==1:
            if ln.endswith('"') and not ln.startswith('"'): ln='"'+ln
            elif ln.startswith('"') and not ln.endswith('"') and ln.count('"')==1: ln=ln+'"'
        lines.append(ln)
    return '\n'.join(lines)
def secinfo(sec):
    s=sec.upper()
    typ='A' if 'TEST A' in s else 'B'
    return typ
def load(tag):
    cfg=CFG[tag]; raw=json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)),f'{tag}_raw.json'))); S=cfg['S']
    out={}
    counters={}
    for n,i in enumerate(raw,1):
        sec=i['sec']; counters[sec]=counters.get(sec,0)+1
        if tag=='m2' and n==81: counters[sec]-=1; continue
        stem=i['stem']
        stem=re.sub(r'^\s*\d+\.\s+','',stem)
        opts=dict(i['opts'])
        for k,v in OPT_FIX.get((tag,n),{}).items(): opts[k]=v
        stem=STEM_FIX.get((tag,n),stem)
        stem=stem.replace('All painted his','Ali painted his')
        if n not in S: continue
        ans,topic,exp=S[n]
        stem=norm(stem) if (tag,n) not in STEM_FIX else stem
        opts={k:norm(v).rstrip('.') if k and (tag,n) in OPT_FIX else norm(v) for k,v in opts.items()}
        mod=[m for m in cfg['modules'] if n in m[3]]
        if not mod: continue
        mod=mod[0]
        m=re.search(r'TEST\s+([AB])',sec.upper())
        label=f"Practice Test {m.group(1)}, câu {counters[sec]}"
        out[n]=dict(id=f'{tag}-{n}',n=n,stem=stem,opts=opts,ans=ans,topic=topic,exp=exp,
                    module=mod[0],mname=mod[1],chapter=mod[2],src=label,excluded=n in cfg['exclude'])
    return out
if __name__=="__main__":
    for tag in CFG:
        d=load(tag)
        print(tag,len(d),sum(1 for x in d.values() if x['excluded']))
        for n,x in d.items():
            if len(x['opts'])!=4 or any(not v for v in x['opts'].values()) or x['ans'] not in x['opts']:
                print("BAD",tag,n,x['opts'])
            if '  ' in x['stem'] or any(len(v)>60 for v in x['opts'].values()): print("LONG",tag,n,x['opts'])
