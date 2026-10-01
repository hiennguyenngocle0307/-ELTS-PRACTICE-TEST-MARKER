import re
def load(path):
    L=[l.strip() for l in open(path) if l.strip()]
    items=[];cur=None;sec=None;i=0
    cur={'src':[],'ans':None};mode='src';
    out=[]
    for l in L:
        m=re.match(r'^Questions\s+(\d+)-(\d+)',l)
        if m: sec=l; continue
        if re.match(r'^[a-d]\.\s',l,re.I) and mode=='src':
            cur['src'].append(l); continue
        if re.match(r'^[a-d]\.\s',l,re.I) and mode=='ans':
            # new item begins
            out.append(cur); cur={'src':[l],'ans':None,'sec':sec}; mode='src'; continue
        if l.lower().startswith('answer'):
            mode='ans'; rest=l.split(':',1)[1].strip() if ':' in l else ''
            cur['ans']=rest; cur['sec']=cur.get('sec') or sec; continue
        if mode=='ans':
            cur['ans']=(cur['ans']+' '+l).strip(); 
    out.append(cur)
    out=[o for o in out if o['src'] and o['ans']]
    return out
if __name__=="__main__":
    S="/tmp/claude-0/s/"
    a=load(S+"Combine_Sentences__MID_GR_-L_N_1.docx.txt")
    b=load(S+"Combine_Sentences__MID_GR_-L_N_2.docx.txt")
    print(len(a),len(b))
    from collections import Counter
    print(Counter(o.get('sec') for o in b))
    A=set(tuple(o['src']) for o in a)
    print("overlap",sum(1 for o in b if tuple(o['src']) in A))
    print(Counter(len(o['src']) for o in a+b))
