import sys,os,re,subprocess,json;sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
import data
from tts import asr
AUD=sys.argv[1]; ONLY=[int(x) for x in sys.argv[2].split(',')] if len(sys.argv)>2 else None
NUM={'0':'zero','1':'one','2':'two','3':'three','4':'four','5':'five','6':'six','7':'seven','8':'eight','9':'nine'}
def canon(s):
    s=s.lower().replace('-',' ').replace(',',' ').replace('.',' ')
    return re.sub(r'[^a-z0-9 ]','',s)
def tokens(s):
    return canon(s).split()
res={}
for t in data.load():
    if ONLY and t['n'] not in ONLY: continue
    for part,rng in ((1,range(1,11)),(2,range(11,21))):
        mp=f"{AUD}/LS1-L-{t['n']:02d}_Part{part}.mp3"; wav=mp[:-4]+'.tmp.wav'
        subprocess.run(['ffmpeg','-y','-loglevel','error','-i',mp,'-ar','16000','-ac','1',wav],check=True)
        txt=asr(wav); os.remove(wav)
        T=canon(txt)
        miss=[]
        for q in rng:
            if part==1:
                ans=t['cues'][q]
                # compare after removing separators; digits/words equivalence
                a=re.sub(r'[^a-z0-9]','',ans.lower()); tt=re.sub(r'[^a-z0-9]','',T)
                # spelled digits
                tt2=tt
                ok=a in tt
                if not ok:
                    aw=''.join(NUM.get(c,c) for c in a); ok= aw in re.sub(r'[^a-z0-9]','',''.join(NUM.get(c,c) for c in T))
                if not ok:
                    # number words e.g. 25 April -> twenty fifth ; accept if letters-only tokens match
                    words=[w for w in re.findall(r'[a-z]+',ans.lower())]
                    ok=bool(words) and all(w in T for w in words) and not re.search(r'\d',ans)
                if not ok: miss.append((q,ans))
        res[(t['n'],part)]=(txt,miss)
        print(t['n'],part,'missing:',miss,flush=True)
        if len(sys.argv)>3: print(txt[:1500])
