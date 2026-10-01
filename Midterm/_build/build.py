# -*- coding: utf-8 -*-
"""Tạo đề giữa kỳ Ngữ pháp (docx + pdf) theo mẫu NP-GK-01.
Dùng:  python3 build.py <thư mục Midterm>   [HTML_OUT=<dir> để xuất HTML cho PDF]"""
import random,re,os,json,sys,html as _h
HERE=os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0,HERE)
from prep import load,CFG
from comb import load as comb_load
from docx import Document
from docx.shared import Pt,Cm,RGBColor,Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT=sys.argv[1] if len(sys.argv)>1 else os.path.join(HERE,'..')
DATA=os.path.join(HERE,'data')+'/'
NCODES=20; PER_MODULE=4
YEAR="Năm học 2026 – 2027 · Học kỳ 1"; COURSE="NGỮ PHÁP TRUNG CẤP"; COURSE_EN="Intermediate Grammar"; COURSE_CODE="711011"; TIME="45 phút"
NAVY="102A43"; PINK="FDE8EB"; LBLUE="E8EEF4"; RED="CC0000"; GREY="5F6368"
BOOK="Azar & Hagen, Understanding and Using English Grammar"
MIDS={1:dict(prefix='NP-GK',title='ĐỀ KIỂM TRA GIỮA KỲ',sub='Đợt 1',
             mods=[('m1',0,'Verb Tenses'),('m1',1,'Singular and Plural'),('m1',2,'Adjective Clauses'),('m1',3,'Noun Clauses'),('m2',3,'Connecting Ideas')]),
      2:dict(prefix='NP-GK2',title='ĐỀ KIỂM TRA GIỮA KỲ',sub='Đợt 2',
             mods=[('m2',0,'Modals and Similar Expressions'),('m2',1,'The Passive'),('m2',2,'Gerunds and Infinitives'),('m2',4,'Showing Relationships Between Ideas'),('m2',5,'Conditional Sentences and Wishes')])}
FILES={'m1':'GRAMMAR_MID_L_N_1','m2':'GRAMMAR_MID_L_N_2'}

# ------------------------------------------------ selection
def combo_pool():
    pool=[];seen=set()
    f1=comb_load(DATA+"Combine_Sentences__MID_GR_-L_N_1.docx.txt")
    for i,o in enumerate(f1):
        k=tuple(o['src'])
        if k in seen: continue
        seen.add(k); pool.append(dict(src=o['src'],ans=o['ans'],ref=f"Combine Sentences – MID GR (Lần 1), Questions 181–210, câu {181+i}"))
    f2=comb_load(DATA+"Combine_Sentences__MID_GR_-L_N_2.docx.txt"); secpos={}
    for o in f2:
        sec=o.get('sec'); secpos[sec]=secpos.get(sec,0)+1
        k=tuple(o['src'])
        if len(k)!=3 or k in seen: continue
        seen.add(k)
        if sec:
            lo=int(re.match(r'Questions\s+(\d+)',sec).group(1))
            ref=f"Combine Sentences – MID GR (Lần 2), Questions {lo}–{lo+29}, câu {lo+secpos[sec]-1}"
        else: ref=f"Combine Sentences – MID GR (Lần 2), phần đầu file (trước nhóm 'Questions 151-180'), mục thứ {secpos[sec]}"
        pool.append(dict(src=o['src'],ans=o['ans'],ref=ref))
    return pool

def letters_plan(rng):
    for _ in range(2000):
        L=list('ABCD')*5; rng.shuffle(L)
        if all(not(L[i]==L[i+1]==L[i+2]) for i in range(18)): return L
    return L
def arrange(q,target,rng):
    correct=q['opts'][q['ans']]; others=[v for k,v in q['opts'].items() if k!=q['ans']]; rng.shuffle(others)
    it=iter(others); return {L:(correct if L==target else next(it)) for L in 'ABCD'}

def make_exams():
    rng=random.Random(20262027)
    banks={t:load(t) for t in CFG}
    pool=combo_pool(); idx=list(range(len(pool))); rng.shuffle(idx)
    exams={1:[],2:[]}; cb_i=0
    for midno in (1,2):
        mods=MIDS[midno]['mods']
        pools=[]
        for (tag,mi,name) in mods:
            rngs=CFG[tag]['modules'][mi][3]
            pools.append([banks[tag][n] for n in rngs if n in banks[tag] and not banks[tag][n]['excluded']])
        use=[{q['id']:0 for q in p} for p in pools]
        for code in range(1,NCODES+1):
            chosen=[]
            for mi,p in enumerate(pools):
                cand=sorted(p,key=lambda q:(use[mi][q['id']],rng.random()))[:PER_MODULE]
                for q in cand: use[mi][q['id']]+=1
                for q in cand: chosen.append((mi,q))
            rng.shuffle(chosen)
            plan=letters_plan(rng); qs=[]
            for k,(mi,q) in enumerate(chosen):
                qs.append(dict(q=q,arr=arrange(q,plan[k],rng),letter=plan[k],mod=mi+1,modname=mods[mi][2],tag=None))
            combs=[pool[i] for i in idx[cb_i:cb_i+4]]; cb_i+=4
            exams[midno].append(dict(code=code,qs=qs,combs=combs))
    return exams

def technique(ans):
    t=[]
    if re.match(r'^[^,]{3,60}, (an? |the )?[^,]{3,}?,',ans) and not re.match(r'^[^,]+, (which|who|that)',ans): t.append('cụm đồng vị (appositive) / cụm danh từ đặt giữa hai dấu phẩy')
    if re.search(r',? which ',ans): t.append('mệnh đề tính từ (adjective clause) với "which"')
    if re.search(r'\bwho\b',ans): t.append('mệnh đề tính từ với "who"')
    if re.search(r'\bthat (is|are|was|were|helps|has|have|it|they)\b',ans) and not t: t.append('mệnh đề tính từ với "that"')
    if re.search(r'\bby \w+ing\b',ans): t.append('giới từ + V-ing (by + V-ing)')
    if re.search(r', (\w+ing|\w+ed|located|known|built|called|used|made|found) ',ans) or re.search(r'^[^,]+, \w+ing\b',ans): t.append('cụm phân từ (adjective phrase)')
    if re.search(r'\b(and|but|yet|so)\b',ans): t.append('cấu trúc song song / liên từ (and, but…)')
    return t or ['kết hợp ý bằng cấu trúc phù hợp']

# ------------------------------------------------ docx helpers
def rgb(h): return RGBColor.from_string(h)
def base_doc():
    d=Document(); s=d.styles['Normal']; s.font.name='Calibri'; s.font.size=Pt(10.5)
    s.element.rPr.rFonts.set(qn('w:eastAsia'),'Calibri')
    s.paragraph_format.space_after=Pt(2); s.paragraph_format.space_before=Pt(0); s.paragraph_format.line_spacing=1.12
    for sec in d.sections:
        sec.page_width=Cm(21.59); sec.page_height=Cm(27.94)
        sec.left_margin=sec.right_margin=Cm(1.8); sec.top_margin=sec.bottom_margin=Cm(1.8)
    return d
def R(p,text,b=False,i=False,sz=None,color=None):
    r=p.add_run(text); r.bold=b; r.italic=i
    if sz: r.font.size=Pt(sz)
    if color: r.font.color.rgb=rgb(color)
    return r
def P(d_or_cell,text='',b=False,i=False,sz=None,color=None,align=None,after=2,before=0,keep=False,indent=None):
    p=d_or_cell.add_paragraph()
    if text: R(p,text,b,i,sz,color)
    if align=='c': p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after=Pt(after); p.paragraph_format.space_before=Pt(before)
    if keep: p.paragraph_format.keep_with_next=True
    if indent is not None: p.paragraph_format.left_indent=Cm(indent)
    return p
def shade_par(p,fill):
    pPr=p._p.get_or_add_pPr(); sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:color'),'auto'); sh.set(qn('w:fill'),fill); pPr.append(sh)
def shade(cell,fill):
    tcPr=cell._tc.get_or_add_tcPr(); sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:color'),'auto'); sh.set(qn('w:fill'),fill); tcPr.append(sh)
def rule(d):
    p=d.add_paragraph(); pPr=p._p.get_or_add_pPr(); b=OxmlElement('w:pBdr'); bt=OxmlElement('w:bottom')
    for k,v in (('val','single'),('sz','6'),('space','6'),('color','C9D6E5')): bt.set(qn('w:'+k),v)
    b.append(bt); pPr.append(b); p.paragraph_format.space_after=Pt(6)
def noborder(t):
    tblPr=t._tbl.tblPr; b=OxmlElement('w:tblBorders')
    for e in ('top','left','bottom','right','insideH','insideV'):
        el=OxmlElement('w:'+e); el.set(qn('w:val'),'nil'); b.append(el)
    tblPr.append(b)
def widths(t,ws):
    t.autofit=False
    for row in t.rows:
        for c,w in zip(row.cells,ws): c.width=Cm(w)
def cellp(cell,text,b=False,i=False,sz=None,color=None,align=None,first=True,after=1):
    p=cell.paragraphs[0] if first else cell.add_paragraph()
    R(p,text,b,i,sz,color); p.paragraph_format.space_after=Pt(after)
    if align=='c': p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    return p
def footer(d,text):
    p=d.sections[0].footer.paragraphs[0]; p.text=''; R(p,text+'  |  Trang ',sz=8.5,color=GREY)
    for t in ('begin','instr','end'):
        rr=p.add_run(); rr.font.size=Pt(8.5)
        if t=='instr':
            it=OxmlElement('w:instrText'); it.set(qn('xml:space'),'preserve'); it.text='PAGE'; rr._r.append(it)
        else:
            fc=OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'),t); rr._r.append(fc)
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
def part_heading(d,left,right):
    p=d.add_paragraph(); shade_par(p,NAVY); p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(6); p.paragraph_format.keep_with_next=True
    R(p,left,b=True,sz=12.5,color='FFFFFF'); R(p,'   '+right,i=True,sz=10.5,color='FFFFFF')

def top_block(d,mid,code,teacher=False):
    cfg=MIDS[mid]
    if teacher: P(d,'TÀI LIỆU GIẢNG VIÊN — KHÔNG PHÁT CHO SINH VIÊN',b=True,sz=9.5,color=RED,align='c',after=3)
    t=d.add_table(rows=1,cols=2); noborder(t); widths(t,[8.2,10])
    l,r=t.rows[0].cells
    cellp(l,'BỘ XÂY DỰNG',sz=10.5,align='c'); cellp(l,'HỌC VIỆN HÀNG KHÔNG VIỆT NAM',b=True,sz=11,align='c',first=False); cellp(l,'KHOA NGOẠI NGỮ',b=True,sz=11,align='c',first=False); cellp(l,'—————',sz=10.5,align='c',first=False)
    ttl=('ĐÁP ÁN VÀ GIẢI THÍCH' if teacher else cfg['title'])+f"  –  {cfg['sub'].upper()}"
    cellp(r,ttl,b=True,sz=16,color=NAVY,align='c'); cellp(r,YEAR,sz=10.5,align='c',first=False)
    cellp(r,'Hình thức: Viết (tự luận + trắc nghiệm)',i=True,sz=10.5,align='c',first=False)
    P(d,'',after=2)
    t=d.add_table(rows=1,cols=2); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER; widths(t,[12.4,5.8])
    l,r=t.rows[0].cells
    cellp(l,f'Học phần: {COURSE}  ({COURSE_EN})',b=True,sz=11); cellp(l,f'Mã học phần: {COURSE_CODE}      Thời gian làm bài: {TIME}',sz=10.5,first=False)
    shade(r,PINK); cellp(r,'MÃ ĐỀ',b=True,sz=10,align='c'); cellp(r,f'{code:02d}',b=True,sz=22,color=RED,align='c',first=False)

def opts_table(d,arr):
    t=d.add_table(rows=2,cols=2); noborder(t); widths(t,[8.9,8.9])
    for (ri,ci),L in zip(((0,0),(0,1),(1,0),(1,1)),'ABCD'):
        c=t.rows[ri].cells[ci]; p=c.paragraphs[0]; R(p,f"{L}.  ",b=True,color=NAVY); R(p,arr[L]); p.paragraph_format.space_after=Pt(1)
        p.paragraph_format.left_indent=Cm(0.5)
    # keep rows together
    for row in t.rows:
        trPr=row._tr.get_or_add_trPr(); cs=OxmlElement('w:cantSplit'); trPr.append(cs)

def grid_letters(d,answers=None,start=1):
    t=d.add_table(rows=6,cols=20); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for blk in range(4):
        for j,h in enumerate(['Câu','A','B','C','D']):
            c=t.rows[0].cells[blk*5+j]; cellp(c,h,b=True,sz=9.5,align='c'); shade(c,LBLUE)
        for r in range(5):
            n=blk*5+r+1
            cellp(t.rows[r+1].cells[blk*5],f'{n}.',b=True,sz=10,align='c')
            for k,L in enumerate('ABCD'):
                c=t.rows[r+1].cells[blk*5+1+k]
                if answers: cellp(c,'●' if answers[n-1]==L else L,b=answers[n-1]==L,sz=10,color=(RED if answers[n-1]==L else GREY),align='c')
                else: cellp(c,L,sz=10,color=GREY,align='c')
    widths(t,[1.0,0.85,0.85,0.85,0.85]*4)
    return t

def build_exam(path,mid,e):
    code=e['code']; d=base_doc(); footer(d,f"{MIDS[mid]['prefix']}-{code:02d}  ·  {COURSE}  ·  Mã đề {code:02d}")
    top_block(d,mid,code)
    p=P(d,'',after=1,before=4); R(p,'Không sử dụng tài liệu [x]      Được sử dụng tài liệu [   ]      Nộp lại đề thi [x]',sz=10)
    P(d,'Đề thi gồm 24 câu (kèm 01 Phiếu trả lời ở trang cuối). Cán bộ coi thi không giải thích gì thêm.',sz=10,after=1)
    P(d,'Sinh viên ghi MÃ ĐỀ và làm bài trên PHIẾU TRẢ LỜI.',b=True,i=True,sz=10,color=NAVY,after=1)
    rule(d)
    part_heading(d,'PART A.  MULTIPLE CHOICE  (Questions 1–20)','8 points · 0.4 point each')
    P(d,'Choose the answer (A, B, C or D) that best completes each sentence. Circle your answers on the ANSWER SHEET.',b=True,i=True,sz=10.5,color=NAVY,after=4)
    for n,x in enumerate(e['qs'],1):
        p=P(d,'',before=5,after=1,keep=True); R(p,f"{n}.  ",b=True,color=NAVY); R(p,' '.join(x['q']['stem'].split('\n')))
        opts_table(d,x['arr'])
    part_heading(d,'PART B.  SENTENCE COMBINATION  (Questions 21–24)','2 points · 0.5 point each')
    P(d,'Combine the three sentences in each item into ONE complete sentence. Keep ALL the information. Use an appositive, an adjective clause / adjective phrase, or parallel structure. Write your answers on the ANSWER SHEET.',b=True,i=True,sz=10.5,color=NAVY,after=3,keep=True)
    p=P(d,'',after=1,keep=True); R(p,'Example:  ',b=True,color=NAVY); R(p,'a. Hanoi is the capital of Vietnam.  b. It is located on the Red River.  c. It has a history of over 1,000 years.',i=True,color=NAVY)
    p=P(d,'',after=4); R(p,'→  Hanoi, the capital of Vietnam, is located on the Red River and has a history of over 1,000 years.',b=True,color=NAVY)
    for n,c in enumerate(e['combs'],21):
        p=P(d,'',before=4,after=1,keep=True); R(p,f"{n}.  ",b=True,color=NAVY)
        for j,ln in enumerate(c['src']):
            q=P(d,'',after=1,keep=(j<2),indent=0.9); m=re.match(r'^([a-d])\.\s*(.*)$',ln); R(q,f"{m.group(1)}.  ",b=True); R(q,m.group(2))
    P(d,'— THE END —',b=True,color=NAVY,align='c',before=10,after=6)
    t=d.add_table(rows=1,cols=2); noborder(t); widths(t,[9,9])
    for c,role in zip(t.rows[0].cells,('Trưởng Bộ môn duyệt','Giảng viên ra đề')):
        cellp(c,'TP. Hồ Chí Minh, ngày ...... tháng ...... năm 2026',i=True,sz=10,align='c'); cellp(c,role,b=True,sz=10.5,align='c',first=False); cellp(c,'(Ký và ghi rõ họ, tên)',i=True,sz=9.5,align='c',first=False)
        for _ in range(3): cellp(c,'',first=False)
    d.add_page_break()
    answer_sheet(d,mid,code)
    d.save(path)

def answer_sheet(d,mid,code):
    P(d,'HỌC VIỆN HÀNG KHÔNG VIỆT NAM  ·  KHOA NGOẠI NGỮ',b=True,sz=10.5,color=NAVY,align='c',after=0)
    P(d,'PHIẾU TRẢ LỜI  /  STUDENT ANSWER SHEET',b=True,sz=17,color=NAVY,align='c',after=0)
    P(d,f'Kiểm tra giữa kỳ ({MIDS[mid]["sub"]}) · {COURSE} · {YEAR}',i=True,sz=10,color=GREY,align='c',after=6)
    t=d.add_table(rows=3,cols=2); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER; widths(t,[13,5.2])
    rows=['Họ và tên:  ..................................................................','MSSV:  ...................................    Lớp:  ..............................','Phòng thi:  ..........    Số tờ:  ..........    Chữ ký SV:  ....................']
    for r_,txt in zip(t.rows,rows): cellp(r_.cells[0],txt,sz=10.5)
    m=t.cell(0,1).merge(t.cell(2,1)); shade(m,PINK); m.text=''; cellp(m,'MÃ ĐỀ',b=True,sz=10,align='c'); cellp(m,f'{code:02d}',b=True,sz=26,color=RED,align='c',first=False)
    P(d,'',after=3)
    t=d.add_table(rows=2,cols=5); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for c,h in zip(t.rows[0].cells,['Part A (/8)','Part B (/2)','TỔNG (/10)','GV chấm 1','GV chấm 2']): cellp(c,h,b=True,sz=9.5,align='c'); shade(c,LBLUE)
    for c,h in zip(t.rows[1].cells,['......... / 8','......... / 2','','','']): cellp(c,h,sz=10,align='c')
    t.rows[1].height=Cm(1.0)
    P(d,'',after=3)
    p=P(d,'',after=3); R(p,'PART A.  MULTIPLE CHOICE',b=True,color=NAVY,sz=11.5); R(p,'   — Circle ONE letter for each question. To change an answer, cross it out (X) and circle another.',i=True,sz=9.5)
    grid_letters(d)
    P(d,'',after=3)
    p=P(d,'',after=3); R(p,'PART B.  SENTENCE COMBINATION',b=True,color=NAVY,sz=11.5); R(p,'   — Write ONE complete sentence for each question.',i=True,sz=9.5)
    for n in range(21,25):
        P(d,f'Question {n}.',b=True,color=NAVY,before=4,after=1)
        for _ in range(2):
            q=P(d,'',after=4); pPr=q._p.get_or_add_pPr(); b=OxmlElement('w:pBdr'); bt=OxmlElement('w:bottom')
            for k,v in (('val','single'),('sz','4'),('space','1'),('color','999999')): bt.set(qn('w:'+k),v)
            b.append(bt); pPr.append(b); q.paragraph_format.space_before=Pt(8)

def build_key(path,mid,e):
    code=e['code']; d=base_doc(); footer(d,f"{MIDS[mid]['prefix']}-{code:02d}  ·  Đáp án (lưu hành nội bộ)")
    top_block(d,mid,code,teacher=True)
    P(d,'',after=2)
    part_heading(d,'PART A.  ĐÁP ÁN NHANH','20 câu · 0.4 điểm/câu')
    grid_letters(d,[x['letter'] for x in e['qs']])
    p=P(d,'',before=4,after=2); R(p,'Phạm vi: ',b=True); R(p,'; '.join(f"Module {i+1} – {m[2]} ({CFG[m[0]]['modules'][m[1]][2]}), 4 câu" for i,m in enumerate(MIDS[mid]['mods']))+'. Thứ tự câu hỏi đã được trộn; đáp án A/B/C/D phân bổ đều (mỗi chữ cái 5 câu).')
    part_heading(d,'PART A.  GIẢI THÍCH CHI TIẾT','')
    for n,x in enumerate(e['qs'],1):
        q=x['q']
        p=P(d,'',before=7,after=1,keep=True); R(p,f'Câu {n}  ',b=True,color=NAVY,sz=11); R(p,f"[Module {x['mod']} · {x['modname']}]",i=True,sz=9.5,color=GREY); R(p,'    Đáp án: ',b=True); R(p,f"{x['letter']}. {x['arr'][x['letter']]}",b=True,color=RED)
        P(d,' '.join(q['stem'].split('\n')).replace('________','_____'),i=True,sz=10,indent=0.6,after=1,keep=True)
        p=P(d,'',indent=0.6,after=1,keep=True); R(p,'Cấu trúc/quy tắc: ',b=True,sz=10); R(p,q['topic']+'. ',sz=10); R(p,q['exp'],sz=10)
        p=P(d,'',indent=0.6,after=2); R(p,'Nguồn: ',b=True,sz=9,color=GREY); R(p,f"{BOOK}, {q['chapter']} – {q['mname']}; ngân hàng câu hỏi {FILES[q['id'].split('-')[0]]} – {q['src']}.",sz=9,color=GREY)
    part_heading(d,'PART B.  ĐÁP ÁN THAM KHẢO & THANG ĐIỂM','2 điểm · 0.5 điểm/câu')
    P(d,'Thang chấm mỗi câu: 0,5 = đúng ngữ pháp, đủ cả 3 ý a–b–c, đúng dấu câu; 0,25 = đủ 3 ý nhưng còn lỗi nhỏ (dấu câu, mạo từ, lặp từ…); 0 = thiếu ý hoặc lỗi cấu trúc nghiêm trọng. Chấp nhận mọi cách nối khác đúng ngữ pháp và đủ thông tin.',i=True,sz=10,after=3)
    for n,c in enumerate(e['combs'],21):
        P(d,f'Câu {n}',b=True,color=NAVY,before=6,after=1,keep=True)
        for ln in c['src']: P(d,ln,sz=10,indent=0.6,after=0,keep=True)
        p=P(d,'',indent=0.6,before=2,after=1,keep=True); R(p,'Đáp án mẫu: ',b=True,sz=10); R(p,c['ans'],b=True,color=RED,sz=10)
        p=P(d,'',indent=0.6,after=1,keep=True); R(p,'Giải thích: ',b=True,sz=10); R(p,'Đáp án mẫu dùng '+'; '.join(technique(c['ans']))+'. Cần giữ đủ thông tin của cả ba câu a, b, c, nối thành một câu duy nhất, tránh lặp chủ ngữ/đại từ (it, they, he…) không cần thiết và dùng dấu câu đúng.',sz=10)
        p=P(d,'',indent=0.6,after=2); R(p,'Nguồn: ',b=True,sz=9,color=GREY); R(p,c['ref']+'.',sz=9,color=GREY)
    d.save(path)

def build_quick(path,mid,exams):
    d=base_doc(); sec=d.sections[0]
    from docx.enum.section import WD_ORIENT
    sec.orientation=WD_ORIENT.LANDSCAPE; sec.page_width,sec.page_height=sec.page_height,sec.page_width
    P(d,'TÀI LIỆU GIẢNG VIÊN — KHÔNG PHÁT CHO SINH VIÊN',b=True,sz=9.5,color=RED,align='c',after=2)
    P(d,'HỌC VIỆN HÀNG KHÔNG VIỆT NAM  ·  KHOA NGOẠI NGỮ',b=True,sz=10.5,color=NAVY,align='c',after=0)
    P(d,f"BẢNG ĐÁP ÁN NHANH — 20 MÃ ĐỀ ({MIDS[mid]['sub'].upper()})",b=True,sz=17,color=NAVY,align='c',after=0)
    P(d,f'Kiểm tra giữa kỳ · {COURSE} ({COURSE_CODE}) · Part A: mỗi câu 0,4đ',i=True,sz=10,color=GREY,align='c',after=6)
    t=d.add_table(rows=21,cols=21); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    for j,h in enumerate(['Mã đề']+[str(i) for i in range(1,21)]):
        c=t.rows[0].cells[j]; cellp(c,h,b=True,sz=9,color='FFFFFF',align='c'); shade(c,NAVY)
    for i,e in enumerate(exams,1):
        cellp(t.rows[i].cells[0],f"{e['code']:02d}",b=True,sz=9.5,align='c'); shade(t.rows[i].cells[0],LBLUE)
        for j,x in enumerate(e['qs'],1): cellp(t.rows[i].cells[j],x['letter'],sz=9.5,align='c')
    widths(t,[1.7]+[1.15]*20)
    P(d,'',after=4)
    P(d,'Part B (câu 21–24): chấm theo đáp án mẫu + thang 0,5 / 0,25 / 0 trong file đáp án của từng mã đề.',sz=10,after=0)
    d.save(path)

# ------------------------------------------------ HTML (PDF)
CSS=f"""<style>@page{{size:Letter;margin:1.8cm}}*{{box-sizing:border-box}}body{{font-family:Calibri,Carlito,'Liberation Sans',Arial,sans-serif;font-size:10.5pt;line-height:1.3;color:#000;margin:0}}
table{{border-collapse:collapse;width:100%}}td,th{{vertical-align:top}}.c{{text-align:center}}.nav{{color:#{NAVY}}}.red{{color:#{RED}}}.gr{{color:#{GREY}}}.b{{font-weight:bold}}.i{{font-style:italic}}
.bx td,.bx th{{border:1px solid #444;padding:4px 6px}}.ph{{background:#{NAVY};color:#fff;padding:5px 8px;margin:10px 0 6px;font-weight:bold;font-size:12.5pt;break-after:avoid}}.ph span{{font-weight:normal;font-style:italic;font-size:10.5pt}}
.q{{margin-top:7px;break-inside:avoid}}.q .st{{margin-bottom:1px}}.og{{display:grid;grid-template-columns:1fr 1fr;gap:0 .4cm;padding-left:.5cm}}.og div{{padding-right:4px}}
.rule{{border-bottom:1px solid #C9D6E5;margin:6px 0 8px}}.k{{margin-left:.6cm}}.sm{{font-size:9pt;color:#{GREY}}}.blk{{break-inside:avoid;margin-top:8px}}.ln{{border-bottom:1px solid #999;height:27px}}
.lt td{{border:1px solid #444;text-align:center;padding:2px 0;font-size:10pt}}.lt th{{border:1px solid #444;background:#{LBLUE};font-size:9.5pt}}.pk{{background:#{PINK}}}</style>"""
def E(t): return _h.escape(t)
def htop(mid,code,teacher=False):
    cfg=MIDS[mid]; ttl=('ĐÁP ÁN VÀ GIẢI THÍCH' if teacher else cfg['title'])+f"  –  {cfg['sub'].upper()}"
    h=("<p class='c b red' style='font-size:9.5pt;margin:0 0 4px'>TÀI LIỆU GIẢNG VIÊN — KHÔNG PHÁT CHO SINH VIÊN</p>" if teacher else "")
    h+=(f"<table><tr><td width=42% class=c>BỘ XÂY DỰNG<br><b>HỌC VIỆN HÀNG KHÔNG VIỆT NAM</b><br><b>KHOA NGOẠI NGỮ</b><br>—————</td>"
        f"<td class=c><div class='b nav' style='font-size:16pt'>{E(ttl)}</div>{YEAR}<br><i>Hình thức: Viết (tự luận + trắc nghiệm)</i></td></tr></table><div style='height:6px'></div>"
        f"<table class=bx><tr><td width=68%><b style='font-size:11pt'>Học phần: {COURSE}  ({COURSE_EN})</b><br>Mã học phần: {COURSE_CODE} &nbsp;&nbsp;&nbsp; Thời gian làm bài: {TIME}</td>"
        f"<td class='c pk'><b style='font-size:10pt'>MÃ ĐỀ</b><br><b class=red style='font-size:22pt;line-height:1.1'>{code:02d}</b></td></tr></table>")
    return h
def hgrid(answers=None):
    h="<table class=lt><tr>"+"".join("<th>Câu</th><th>A</th><th>B</th><th>C</th><th>D</th>" for _ in range(4))+"</tr>"
    for r in range(5):
        h+="<tr>"
        for blk in range(4):
            n=blk*5+r+1; h+=f"<td><b>{n}.</b></td>"
            for L in 'ABCD':
                if answers: h+=("<td class='b red'>●</td>" if answers[n-1]==L else f"<td class=gr>{L}</td>")
                else: h+=f"<td class=gr>{L}</td>"
        h+="</tr>"
    return h+"</table>"
def exam_html(mid,e):
    code=e['code']; h=f"<html><head><meta charset=utf-8>{CSS}</head><body>"+htop(mid,code)
    h+=("<p style='margin:8px 0 2px;font-size:10pt'>Không sử dụng tài liệu [x] &nbsp;&nbsp;&nbsp; Được sử dụng tài liệu [&nbsp;&nbsp;&nbsp;] &nbsp;&nbsp;&nbsp; Nộp lại đề thi [x]</p>"
        "<p style='margin:0 0 2px;font-size:10pt'>Đề thi gồm 24 câu (kèm 01 Phiếu trả lời ở trang cuối). Cán bộ coi thi không giải thích gì thêm.</p>"
        "<p class='b i nav' style='margin:0;font-size:10pt'>Sinh viên ghi MÃ ĐỀ và làm bài trên PHIẾU TRẢ LỜI.</p><div class=rule></div>")
    h+="<div class=ph>PART A.  MULTIPLE CHOICE  (Questions 1–20) <span>&nbsp; 8 points · 0.4 point each</span></div><p class='b i nav' style='margin:0 0 4px'>Choose the answer (A, B, C or D) that best completes each sentence. Circle your answers on the ANSWER SHEET.</p>"
    for n,x in enumerate(e['qs'],1):
        h+=f"<div class=q><div class=st><b class=nav>{n}.</b>&nbsp; {E(' '.join(x['q']['stem'].split(chr(10))))}</div><div class=og>"+"".join(f"<div><b class=nav>{L}.</b>&nbsp; {E(x['arr'][L])}</div>" for L in 'ABCD')+"</div></div>"
    h+="<div style='break-inside:avoid'><div class=ph>PART B.  SENTENCE COMBINATION  (Questions 21–24) <span>&nbsp; 2 points · 0.5 point each</span></div>"
    h+="<p class='b i nav' style='margin:0 0 3px'>Combine the three sentences in each item into ONE complete sentence. Keep ALL the information. Use an appositive, an adjective clause / adjective phrase, or parallel structure. Write your answers on the ANSWER SHEET.</p>"
    h+="<p class=nav style='margin:0'><b>Example:</b> <i>a. Hanoi is the capital of Vietnam.  b. It is located on the Red River.  c. It has a history of over 1,000 years.</i></p><p class='b nav' style='margin:0 0 4px'>→ Hanoi, the capital of Vietnam, is located on the Red River and has a history of over 1,000 years.</p></div>"
    for n,c in enumerate(e['combs'],21):
        h+=f"<div class=q><b class=nav>{n}.</b>"+"".join(f"<div style='margin-left:.9cm'><b>{E(ln[0])}.</b>&nbsp; {E(ln[2:].strip())}</div>" for ln in c['src'])+"</div>"
    h+="<p class='c b nav' style='margin:12px 0 6px'>— THE END —</p><table style='break-inside:avoid'><tr>"+"".join(f"<td class=c style='font-size:10pt'><i>TP. Hồ Chí Minh, ngày ...... tháng ...... năm 2026</i><br><b>{r}</b><br><i style='font-size:9.5pt'>(Ký và ghi rõ họ, tên)</i><div style='height:50px'></div></td>" for r in ('Trưởng Bộ môn duyệt','Giảng viên ra đề'))+"</tr></table>"
    h+="<div style='break-before:page'>"
    h+=(f"<p class='c b nav' style='margin:0'>HỌC VIỆN HÀNG KHÔNG VIỆT NAM &nbsp;·&nbsp; KHOA NGOẠI NGỮ</p><p class='c b nav' style='margin:0;font-size:17pt'>PHIẾU TRẢ LỜI &nbsp;/&nbsp; STUDENT ANSWER SHEET</p><p class='c i gr' style='margin:0 0 8px;font-size:10pt'>Kiểm tra giữa kỳ ({MIDS[mid]['sub']}) · {COURSE} · {YEAR}</p>"
        f"<table class=bx><tr><td width=72%>Họ và tên: ..................................................................</td><td rowspan=3 class='c pk'><b style='font-size:10pt'>MÃ ĐỀ</b><br><b class=red style='font-size:26pt'>{code:02d}</b></td></tr>"
        "<tr><td>MSSV: ................................... &nbsp;&nbsp; Lớp: ..............................</td></tr><tr><td>Phòng thi: .......... &nbsp;&nbsp; Số tờ: .......... &nbsp;&nbsp; Chữ ký SV: ....................</td></tr></table><div style='height:8px'></div>"
        "<table class='bx c'><tr style='background:#"+LBLUE+"'><th>Part A (/8)</th><th>Part B (/2)</th><th>TỔNG (/10)</th><th>GV chấm 1</th><th>GV chấm 2</th></tr><tr style='height:32px'><td>......... / 8</td><td>......... / 2</td><td></td><td></td><td></td></tr></table><div style='height:10px'></div>")
    h+="<p style='margin:0 0 4px'><b class=nav style='font-size:11.5pt'>PART A.  MULTIPLE CHOICE</b> &nbsp;<i style='font-size:9.5pt'>— Circle ONE letter for each question. To change an answer, cross it out (X) and circle another.</i></p>"+hgrid()
    h+="<p style='margin:12px 0 0'><b class=nav style='font-size:11.5pt'>PART B.  SENTENCE COMBINATION</b> &nbsp;<i style='font-size:9.5pt'>— Write ONE complete sentence for each question.</i></p>"
    for n in range(21,25): h+=f"<p class='b nav' style='margin:8px 0 0'>Question {n}.</p><div class=ln></div><div class=ln></div>"
    return h+"</div></body></html>"
def key_html(mid,e):
    code=e['code']; h=f"<html><head><meta charset=utf-8>{CSS}</head><body>"+htop(mid,code,True)
    h+="<div class=ph>PART A.  ĐÁP ÁN NHANH <span>&nbsp; 20 câu · 0.4 điểm/câu</span></div>"+hgrid([x['letter'] for x in e['qs']])
    h+="<p style='margin:6px 0'><b>Phạm vi:</b> "+E('; '.join(f"Module {i+1} – {m[2]} ({CFG[m[0]]['modules'][m[1]][2]}), 4 câu" for i,m in enumerate(MIDS[mid]['mods'])))+". Thứ tự câu hỏi đã được trộn; đáp án A/B/C/D phân bổ đều (mỗi chữ cái 5 câu).</p>"
    h+="<div class=ph>PART A.  GIẢI THÍCH CHI TIẾT</div>"
    for n,x in enumerate(e['qs'],1):
        q=x['q']
        h+=(f"<div class=blk><b class=nav style='font-size:11pt'>Câu {n}</b> <i class=gr style='font-size:9.5pt'>[Module {x['mod']} · {E(x['modname'])}]</i> &nbsp;<b>Đáp án: <span class=red>{x['letter']}. {E(x['arr'][x['letter']])}</span></b>"
            f"<div class=k style='font-size:10pt'><i>{E(' '.join(q['stem'].split(chr(10))).replace('________','_____'))}</i></div>"
            f"<div class=k style='font-size:10pt'><b>Cấu trúc/quy tắc:</b> {E(q['topic'])}. {E(q['exp'])}</div>"
            f"<div class='k sm'><b>Nguồn:</b> {E(BOOK)}, {E(q['chapter'])} – {E(q['mname'])}; ngân hàng câu hỏi {FILES[q['id'].split('-')[0]]} – {E(q['src'])}.</div></div>")
    h+="<div class=ph>PART B.  ĐÁP ÁN THAM KHẢO & THANG ĐIỂM <span>&nbsp; 2 điểm · 0.5 điểm/câu</span></div><p class=i style='font-size:10pt'>Thang chấm mỗi câu: 0,5 = đúng ngữ pháp, đủ cả 3 ý a–b–c, đúng dấu câu; 0,25 = đủ 3 ý nhưng còn lỗi nhỏ (dấu câu, mạo từ, lặp từ…); 0 = thiếu ý hoặc lỗi cấu trúc nghiêm trọng. Chấp nhận mọi cách nối khác đúng ngữ pháp và đủ thông tin.</p>"
    for n,c in enumerate(e['combs'],21):
        h+=f"<div class=blk><b class=nav>Câu {n}</b>"+"".join(f"<div class=k style='font-size:10pt'>{E(l)}</div>" for l in c['src'])
        h+=(f"<div class=k style='font-size:10pt'><b>Đáp án mẫu:</b> <b class=red>{E(c['ans'])}</b></div><div class=k style='font-size:10pt'><b>Giải thích:</b> Đáp án mẫu dùng {E('; '.join(technique(c['ans'])))}. Cần giữ đủ thông tin của cả ba câu a, b, c, nối thành một câu duy nhất, tránh lặp chủ ngữ/đại từ (it, they, he…) không cần thiết và dùng dấu câu đúng.</div>"
            f"<div class='k sm'><b>Nguồn:</b> {E(c['ref'])}.</div></div>")
    return h+"</body></html>"
def quick_html(mid,exams):
    h=f"<html><head><meta charset=utf-8>{CSS.replace('size:Letter','size:Letter landscape')}</head><body>"
    h+=(f"<p class='c b red' style='font-size:9.5pt;margin:0'>TÀI LIỆU GIẢNG VIÊN — KHÔNG PHÁT CHO SINH VIÊN</p><p class='c b nav' style='margin:0'>HỌC VIỆN HÀNG KHÔNG VIỆT NAM &nbsp;·&nbsp; KHOA NGOẠI NGỮ</p>"
        f"<p class='c b nav' style='margin:0;font-size:17pt'>BẢNG ĐÁP ÁN NHANH — 20 MÃ ĐỀ ({MIDS[mid]['sub'].upper()})</p><p class='c i gr' style='margin:0 0 8px;font-size:10pt'>Kiểm tra giữa kỳ · {COURSE} ({COURSE_CODE}) · Part A: mỗi câu 0,4đ</p>")
    h+="<table class=lt><tr><th style='background:#"+NAVY+";color:#fff'>Mã đề</th>"+"".join(f"<th style='background:#{NAVY};color:#fff'>{i}</th>" for i in range(1,21))+"</tr>"
    for e in exams: h+=f"<tr><td class=b style='background:#{LBLUE}'>{e['code']:02d}</td>"+"".join(f"<td>{x['letter']}</td>" for x in e['qs'])+"</tr>"
    return h+"</table><p style='font-size:10pt'>Part B (câu 21–24): chấm theo đáp án mẫu + thang 0,5 / 0,25 / 0 trong file đáp án của từng mã đề.</p></body></html>"

# ------------------------------------------------ main
def main():
    exams=make_exams(); comp=[]
    root=os.path.join(OUT,'Grammar','TÂM LINH Midterm Test'); hout=os.environ.get('HTML_OUT')
    for mid in (1,2):
        base=os.path.join(root,f'Midterm_{mid}'); dt=os.path.join(base,'De_thi_va_Phieu_tra_loi'); da=os.path.join(base,'Dap_an')
        os.makedirs(dt,exist_ok=True); os.makedirs(da,exist_ok=True)
        pre=MIDS[mid]['prefix']
        for e in exams[mid]:
            nm=f"{pre}-{e['code']:02d}"
            build_exam(os.path.join(dt,f'{nm} - DE THI.docx'),mid,e)
            build_key(os.path.join(da,f'{nm} - DAP AN.docx'),mid,e)
            comp.append(dict(mid=mid,code=e['code'],mc=[x['q']['id'] for x in e['qs']],ans=''.join(x['letter'] for x in e['qs']),comb=[c['ref'] for c in e['combs']]))
            if hout:
                for sub,fn,hh in (('De_thi_va_Phieu_tra_loi',f'{nm} - DE THI',exam_html(mid,e)),('Dap_an',f'{nm} - DAP AN',key_html(mid,e))):
                    pd=os.path.join(hout,f'Midterm_{mid}',sub); os.makedirs(pd,exist_ok=True); open(os.path.join(pd,fn+'.html'),'w',encoding='utf-8').write(hh)
        build_quick(os.path.join(da,'00 - BANG DAP AN NHANH 20 MA DE.docx'),mid,exams[mid])
        if hout:
            pd=os.path.join(hout,f'Midterm_{mid}','Dap_an'); open(os.path.join(pd,'00 - BANG DAP AN NHANH 20 MA DE.html'),'w',encoding='utf-8').write(quick_html(mid,exams[mid]))
    json.dump(comp,open(os.path.join(HERE,'exam_composition.json'),'w'),ensure_ascii=False,indent=1)
if __name__=="__main__": main()
