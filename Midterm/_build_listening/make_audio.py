# -*- coding: utf-8 -*-
"""Tạo 2 file ghi âm (Part 1, Part 2) cho mỗi test từ transcript nguồn bằng TTS offline (Piper qua sherpa-onnx)."""
import re,os,sys,json,subprocess,numpy as np
HERE=os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0,HERE)
import data
from tts import synth,write_wav
OUT=sys.argv[1] if len(sys.argv)>1 else os.path.join(HERE,'audio_out')
ONLY=[int(x) for x in sys.argv[2].split(',')] if len(sys.argv)>2 else None
SR=22050
W2={'0':'zero','1':'one','2':'two','3':'three','4':'four','5':'five','6':'six','7':'seven','8':'eight','9':'nine'}
def digits(s): return ' '.join(W2[c] for c in s if c in W2)
def year(n):
    a,b=n[:2],n[2:]
    tw={'10':'ten','11':'eleven','12':'twelve','13':'thirteen','14':'fourteen','15':'fifteen','16':'sixteen','17':'seventeen','18':'eighteen','19':'nineteen','20':'twenty'}
    ones=['','one','two','three','four','five','six','seven','eight','nine']; tens={'2':'twenty','3':'thirty','4':'forty','5':'fifty','6':'sixty','7':'seventy','8':'eighty','9':'ninety'}
    if b=='00': return tw[a]+' hundred'
    if b[0]=='0': return tw[a]+' oh '+ones[int(b[1])]
    if b[0]=='1': return tw[a]+' '+{'10':'ten','11':'eleven','12':'twelve','13':'thirteen','14':'fourteen','15':'fifteen','16':'sixteen','17':'seventeen','18':'eighteen','19':'nineteen'}[b]
    return tw[a]+' '+tens[b[0]]+('-'+ones[int(b[1])] if b[1]!='0' else '')
def norm(t):
    t=re.sub(r'\[Q\d+: [^\]]*\]','',t)
    t=re.sub(r'\b(twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety)-(one|two|three|four|five|six|seven|eight|nine)\b',r'\1 \2',t)
    t=t.replace('—',', ').replace('–',', ').replace('’',"'")
    t=re.sub(r'\b([A-Z](?:-[A-Z]){2,})\b',lambda m:', '.join(m.group(1).split('-'))+',',t)
    t=t.replace('m.jones at greenmail dot com','M dot Jones at greenmail dot com')
    t=t.replace('LS8 4RP','L S eight, four R P')
    t=re.sub(r'\b(HB|JC|CR|MV) (?=\w)',lambda m:' '.join(m.group(1))+' ',t)
    t=re.sub(r'\b0(\d{2,4}) (\d{3,6})(?: (\d{3,4}))?\b',lambda m:'oh '+digits(m.group(1))+', '+digits(m.group(2))+((', '+digits(m.group(3))) if m.group(3) else ''),t)
    t=re.sub(r'\b(1[5-9]\d\d|20\d\d)\b',lambda m:year(m.group(1)),t)
    t=re.sub(r'\b(Room|Zone|Studio|floor|Floor) (\d+)\b',lambda m:m.group(1)+' '+digits(m.group(2)),t)
    t=re.sub(r'\b(\d+)\b',lambda m:digits(m.group(1)),t)
    t=re.sub(r'\s+',' ',t).strip()
    return t
def sil(sec): return np.zeros(int(SR*sec),dtype=np.float32)
def speak(text,voice,speed=1.0):
    parts=[p for p in re.split(r'(?<=[.?!])\s+',text) if p.strip()]
    out=[]
    for p in parts:
        x,sr=synth(p,voice,speed); assert sr==SR,sr
        out+= [x,sil(0.18)]
    return np.concatenate(out)
PAIRS=[('alan','jenny'),('alan','alba'),('jenny','alba'),('alba','alan')]
def build(t):
    n=t['n']; d1,d2=PAIRS[(n-1)%4]
    ann=[v for v in ('alan','jenny','alba') if v not in (d1,d2)][0]
    mono=(d1,d2)[n%2]
    # ---- Part 1
    turns=[]
    for l in t['script1']:
        m=re.match(r'^([A-Za-z ]+):\s*(.*)$',l); turns.append((m.group(1),norm(m.group(2))))
    spk=[]; 
    for s,_ in turns:
        if s not in spk: spk.append(s)
    vmap={spk[0]:d1,spk[1]:d2}
    a=[]
    a.append(speak(f"Listening and Speaking One. Midterm test. Test {n}. Part 1. Questions 1 to 10. You will hear a conversation between two people. Listen and complete the form. You will hear the recording once only. You now have twenty seconds to look at Part 1.",ann)); a.append(sil(20))
    a.append(speak("Now listen to the recording.",ann)); a.append(sil(1.2))
    for s,tx in turns: a.append(speak(tx,vmap[s])); a.append(sil(0.55))
    a.append(sil(1.0)); a.append(speak("That is the end of Part 1. You now have thirty seconds to check your answers.",ann)); a.append(sil(30)); a.append(speak("This is the end of the recording for Part 1.",ann))
    p1=np.concatenate(a)
    # ---- Part 2
    pa=norm(t['script2'][0]); pb=norm(t['script2'][1])
    b=[]
    b.append(speak(f"Listening and Speaking One. Midterm test. Test {n}. Part 2. Questions 11 to 20. You will hear a talk. You will hear the recording once only. First, you have twenty seconds to look at questions 11 to 15.",ann)); b.append(sil(20))
    b.append(speak("Now listen and answer questions 11 to 15.",ann)); b.append(sil(1.2))
    b.append(speak(pa,mono)); b.append(sil(1.2))
    b.append(speak("Before you hear the rest of the talk, you have fifteen seconds to look at questions 16 to 20.",ann)); b.append(sil(15))
    b.append(speak("Now listen and answer questions 16 to 20.",ann)); b.append(sil(1.2))
    b.append(speak(pb,mono)); b.append(sil(1.0))
    b.append(speak("That is the end of Part 2. You now have thirty seconds to check your answers.",ann)); b.append(sil(30)); b.append(speak("This is the end of the recording for Part 2.",ann))
    p2=np.concatenate(b)
    return p1,p2
def mp3(x,path):
    wav=path[:-4]+'.wav'; write_wav(wav,x,SR)
    subprocess.run(['ffmpeg','-y','-loglevel','error','-i',wav,'-ac','1','-b:a','56k','-ar','22050',path],check=True); os.remove(wav)
if __name__=="__main__":
    os.makedirs(OUT,exist_ok=True); dur={}
    for t in data.load():
        if ONLY and t['n'] not in ONLY: continue
        p1,p2=build(t)
        mp3(p1,f"{OUT}/LS1-L-{t['n']:02d}_Part1.mp3"); mp3(p2,f"{OUT}/LS1-L-{t['n']:02d}_Part2.mp3")
        dur[t['n']]=(len(p1)/SR,len(p2)/SR); print(t['n'],[round(len(p)/SR/60,1) for p in (p1,p2)],flush=True)
    json.dump(dur,open(os.path.join(OUT,'durations.json') if not ONLY else os.path.join(OUT,'durations_part.json'),'w'))
