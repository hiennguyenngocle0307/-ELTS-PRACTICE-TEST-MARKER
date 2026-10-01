import sys,os,re,json;sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
import numpy as np,sherpa_onnx
import data,make_audio as M,tts
d="/tmp/claude-0/tts/sherpa-onnx-whisper-small.en"
rec=sherpa_onnx.OfflineRecognizer.from_whisper(encoder=f"{d}/small.en-encoder.int8.onnx",decoder=f"{d}/small.en-decoder.int8.onnx",tokens=f"{d}/small.en-tokens.txt",language='en',task='transcribe',num_threads=4)
def transcribe(x,sr):
    s=rec.create_stream(); s.accept_waveform(sr,x); rec.decode_stream(s); return s.result.text.strip()
only=[int(a) for a in sys.argv[1].split(',')] if len(sys.argv)>1 else None
for t in data.load():
    if only and t['n'] not in only: continue
    n=t['n']; d1,d2=M.PAIRS[(n-1)%4]; turns=[]
    for l in t['script1']:
        m=re.match(r'^([A-Za-z ]+):\s*(.*)$',l); turns.append((m.group(1),l,M.norm(m.group(2))))
    spk=[]
    for s,_,_ in turns:
        if s not in spk: spk.append(s)
    vm={spk[0]:d1,spk[1]:d2}
    for s,raw,tx in turns:
        if '[Q' not in raw: continue
        x=M.speak(tx,vm[s]); out=transcribe(np.concatenate([M.sil(0.3),x,M.sil(0.3)]),M.SR)
        q=[int(a) for a in re.findall(r'\[Q(\d+):',raw)][0]
        print(f"T{n} Q{q} key={t['key'][q]!r:22} asr={out}",flush=True)
