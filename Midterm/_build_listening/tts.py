import sherpa_onnx,numpy as np,os,re,wave
M=os.environ.get('TTS_DIR','/tmp/claude-0/tts')
VOICES={'alan':'vits-piper-en_GB-alan-medium','jenny':'vits-piper-en_GB-jenny_dioco-medium','north':'vits-piper-en_GB-northern_english_male-medium','alba':'vits-piper-en_GB-alba-medium'}
_cache={}
def engine(name):
    if name in _cache: return _cache[name]
    d=f"{M}/{VOICES[name]}"; onnx=[f for f in os.listdir(d) if f.endswith('.onnx')][0]
    cfg=sherpa_onnx.OfflineTtsConfig(model=sherpa_onnx.OfflineTtsModelConfig(vits=sherpa_onnx.OfflineTtsVitsModelConfig(model=f"{d}/{onnx}",tokens=f"{d}/tokens.txt",data_dir=f"{d}/espeak-ng-data",noise_scale=0.6,noise_scale_w=0.7,length_scale=1.08),num_threads=4,provider='cpu'))
    _cache[name]=sherpa_onnx.OfflineTts(cfg); return _cache[name]
def synth(text,voice,speed=1.0):
    g=engine(voice).generate(text,sid=0,speed=speed)
    return np.array(g.samples,dtype=np.float32),g.sample_rate
def write_wav(path,x,sr):
    with wave.open(path,'wb') as w:
        w.setnchannels(1);w.setsampwidth(2);w.setframerate(sr);w.writeframes((np.clip(x,-1,1)*32767).astype(np.int16).tobytes())
def asr(path):
    d=f"{M}/sherpa-onnx-whisper-tiny.en"
    r=sherpa_onnx.OfflineRecognizer.from_whisper(encoder=f"{d}/tiny.en-encoder.int8.onnx",decoder=f"{d}/tiny.en-decoder.int8.onnx",tokens=f"{d}/tiny.en-tokens.txt",language='en',task='transcribe',num_threads=4)
    import wave as W
    w=W.open(path); sr=w.getframerate(); x=np.frombuffer(w.readframes(w.getnframes()),dtype=np.int16).astype(np.float32)/32768
    out=[]; step=sr*20
    for i in range(0,len(x),step):
        s=r.create_stream(); s.accept_waveform(sr,x[i:i+sr*28]); r.decode_stream(s); out.append(s.result.text.strip())
    return ' '.join(out)
