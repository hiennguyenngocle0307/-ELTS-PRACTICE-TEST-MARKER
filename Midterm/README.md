# Midterm – Học viện Hàng không Vũ trụ Việt Nam (Vietnam Aerospace University)

Bộ đề thi giữa kỳ, đặt tên **Midterm-TÂM LINH**.

```
Midterm/
└── Grammar/
    └── TÂM LINH Midterm Test/          (thư mục riêng của bộ đề Midterm-TÂM LINH)
        ├── Midterm_1/                  (đợt 1 – nguồn: GRAMMAR_MID_L_N_1 + Combine Sentences MID GR Lần 1)
        │   ├── De_thi_va_Phieu_tra_loi/    20 đề (.docx + .pdf): đề thi + Student Answer Sheet (trang cuối)
        │   └── Dap_an/                     20 đáp án (.docx + .pdf): Answer Key + giải thích chi tiết + nguồn (giáo viên)
        └── Midterm_2/                  (đợt 2 – nguồn: GRAMMAR_MID_L_N_2 + Combine Sentences MID GR Lần 2)
            ├── De_thi_va_Phieu_tra_loi/
            └── Dap_an/
```

## Cấu trúc mỗi mã đề (24 câu, 10 điểm, 45 phút)
- **Part I – 20 trắc nghiệm** (0,4 đ/câu): 5 câu/module, bốc ngẫu nhiên có cân bằng từ ngân hàng câu hỏi; vị trí đáp án A/B/C/D được xáo trộn (mỗi đề đúng 5 A, 5 B, 5 C, 5 D).
- **Part II – 4 câu nối câu** (0,5 đ/câu): gộp 3 câu a–b–c thành 1 câu.
- Mỗi đề có **Student Answer Sheet** ở trang cuối; mỗi đáp án có phần giải thích (cấu trúc/quy tắc, lý do loại phương án sai) và **nguồn trích dẫn**.

| | Module 1 | Module 2 | Module 3 | Module 4 |
|---|---|---|---|---|
| **Midterm 1** (câu 1–40 / 41–80 / 81–120 / 121–160) | Verb Tenses (Ch.1) | Singular & Plural (Ch.5) | Adjective Clauses (Ch.6) | Noun Clauses (Ch.7) |
| **Midterm 2** | Modals (Ch.2) | Passive (Ch.3) | Gerunds & Infinitives (Ch.4) | Connecting Ideas – adverb clauses (Ch.8) |

## Những điểm cần giáo viên xác nhận
1. **20 đề Grammar mẫu và "Phụ lục 2" không có trong file đính kèm** (chỉ nhận 4 file .docx). Format/style dựa trên chính các file ngân hàng; nếu có đề mẫu, gửi lại để chỉnh header, thời gian, thang điểm.
2. **Midterm 2 – Module 4**: file nguồn có Chapter 8, 9, 10; đề dùng 4 chương đầu theo thứ tự file (Ch.2, 3, 4, 8). Đổi sang Ch.9/Ch.10 chỉ cần sửa cấu hình trong `_build/prep.py`.
3. **Nối câu Midterm 1**: file nguồn chỉ có **30 mục** (Questions 181–210) nhưng cần 80 vị trí (20 đề × 4) → mỗi mục xuất hiện 2–3 lần ở các đề khác nhau (không trùng trong cùng một đề). Midterm 2 dùng 80 mục khác nhau, đã loại nhóm "181-210 (MID 1 2026-2027)" và các mục trùng với Midterm 1.
4. Đáp án trắc nghiệm do **tự giải** và đối chiếu với các câu đã highlight vàng trong file nguồn (khớp 100% ở các câu có highlight, trừ m2-92: file đánh dấu D "watching to land it" – sai ngữ pháp; đề dùng B "watching it land").
5. **Nguồn trích dẫn**: ghi chương/chủ điểm của *Understanding and Using English Grammar* (Azar & Hagen) và vị trí câu trong ngân hàng gốc; **chưa có số trang/số Chart** vì không có sách giáo trình.
6. **Câu bị loại khỏi ngân hàng do lỗi nguồn** (trùng đáp án, thiếu đề, hai đáp án đúng): Midterm 1 – m1-102, 119, 157; Midterm 2 – m2-2, 89, 123, 124, 133, 139. Một số lỗi chính tả/đáp án trùng lặp ở phương án nhiễu đã được sửa nhẹ (vd: m1-43, 44, 73; m2-64, 120).
7. PDF đã xem thử hiển thị; file .docx chưa được kiểm tra bằng hiển thị (môi trường không chạy được LibreOffice) – nên mở thử một vài mã đề trong Word trước khi in.

## Định dạng
Mỗi đề và mỗi đáp án có cả **.docx** (chỉnh sửa được) và **.pdf** (in ấn) cùng tên, cùng thư mục. PDF được xuất từ HTML (`_build/html2pdf.js`), nội dung giống file Word, bố cục có thể khác đôi chút.

## Tạo lại / chỉnh sửa
`_build/` chứa mã nguồn tạo đề: `cd Midterm/_build && python3 build.py ..` (cần `python-docx`). `exam_composition.json` ghi câu hỏi và đáp án của từng mã đề.
