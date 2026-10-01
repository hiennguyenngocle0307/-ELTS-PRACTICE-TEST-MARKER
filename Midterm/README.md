# Midterm Grammar – NGỮ PHÁP TRUNG CẤP (711011) – bộ đề TÂM LINH

Định dạng theo file mẫu `NP-GK-01 - DE THI` và `00 - BANG DAP AN NHANH 20 MA DE`.

```
Midterm/Grammar/TÂM LINH Midterm Test/
├── Midterm_1/   (đợt 1: mã đề NP-GK-01 … NP-GK-20)
│   ├── De_thi_va_Phieu_tra_loi/   "NP-GK-01 - DE THI" .docx + .pdf (đề + Phiếu trả lời ở trang cuối)
│   └── Dap_an/                    "NP-GK-01 - DAP AN" .docx + .pdf (đáp án + giải thích + nguồn)
│                                  "00 - BANG DAP AN NHANH 20 MA DE" .docx + .pdf
└── Midterm_2/   (đợt 2: mã đề NP-GK2-01 … NP-GK2-20, cùng cấu trúc)
```

## Cấu trúc mỗi mã đề (24 câu, 10 điểm, 45 phút)
- **Part A** – 20 câu trắc nghiệm (0,4 đ/câu) = 5 nhóm × 4 câu, thứ tự câu đã trộn, đáp án A/B/C/D phân bổ đều (5 câu mỗi chữ).
- **Part B** – 4 câu nối câu (0,5 đ/câu; chấm 0,5 / 0,25 / 0), có ví dụ mẫu.
- Mỗi nhóm 4 câu lấy từ: **Midterm 1** – Ch.1 Verb Tenses, Ch.5 Singular & Plural, Ch.6 Adjective Clauses, Ch.7 Noun Clauses (file GRAMMAR_MID_L_N_1) + Ch.8 Connecting Ideas (file GRAMMAR_MID_L_N_2); **Midterm 2** – Ch.2 Modals, Ch.3 Passive, Ch.4 Gerunds & Infinitives, Ch.9 Showing Relationships Between Ideas, Ch.10 Conditional Sentences & Wishes (file GRAMMAR_MID_L_N_2).
- 160 mục nối câu khác nhau cho 40 mã đề (không lặp), lấy từ hai file Combine Sentences.

## Cần giáo viên xác nhận / lưu ý
1. Cách chia 5 nhóm × 4 câu và Ch.8 thuộc Midterm 1 được suy ra từ đề mẫu NP-GK-01 (đề mẫu có 4 câu Verb Tenses, 4 Singular/Plural, 4 Adjective Clauses, 4 Noun Clauses, 4 Connecting Ideas). Midterm 2 dùng các chương còn lại (2, 3, 4, 9, 10) – sửa trong `_build/build.py` (`MIDS`) nếu cần.
2. Mẫu ghi "04 trang giấy"; bản này ghi "24 câu (kèm 01 Phiếu trả lời ở trang cuối)" vì số trang phụ thuộc cách in.
3. Đáp án trắc nghiệm do tự giải, đối chiếu với các câu highlight vàng trong file nguồn (khớp 100%, trừ m2-92 – file đánh dấu sai; đề đã loại hoặc dùng đáp án đúng ngữ pháp).
4. Nguồn trích dẫn: chương/chủ điểm của *Understanding and Using English Grammar* (Azar & Hagen) + vị trí câu trong ngân hàng gốc; chưa có số trang/Chart vì không có sách.
5. Câu bị loại do lỗi nguồn: m1-102, 119, 157; m2-2, 89, 123, 124, 133, 139, 193, 209. Một số phương án nhiễu gõ sai/trùng đã được sửa nhẹ.
6. File .docx chưa kiểm tra bằng hiển thị (môi trường không chạy được LibreOffice); PDF đã xem thử, bố cục PDF và Word có thể khác chút – nên mở thử vài mã đề trong Word trước khi in.

## Tạo lại
`cd Midterm/_build && python3 build.py ..` (cần `python-docx`). PDF: đặt `HTML_OUT=<thư mục>` khi chạy `build.py`, rồi `node html2pdf.js <HTML_OUT> <thư mục PDF>` (cần playwright). `exam_composition.json` ghi thành phần từng mã đề.
