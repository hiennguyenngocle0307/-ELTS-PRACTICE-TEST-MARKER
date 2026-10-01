# LS1-MIDTERM TEST – LISTENING (Nghe – Nói 1, phần Listening)

12 mã đề (LS1-L-01 … LS1-L-12), mỗi đề có **2 file ghi âm** (Part 1 và Part 2), đề thi + phiếu trả lời, đáp án giáo viên có giải thích. Trình bày theo mẫu bộ đề Grammar (header Học viện, ô MÃ ĐỀ, Part heading màu xanh navy, phiếu trả lời cuối đề).

```
LS1-MIDTERM TEST - LISTENING/
├── 00 - PLAY RECORDINGS.html          mở bằng trình duyệt để nghe cả 24 file
├── Audio/                             LS1-L-01_Part1.mp3, LS1-L-01_Part2.mp3, …
├── De_thi_va_Phieu_tra_loi/           "LS1-L-01 - DE THI" .docx + .pdf   (có link mở file ghi âm)
└── Dap_an/                            "LS1-L-01 - DAP AN" .docx + .pdf  +  "00 - BANG DAP AN NHANH 12 MA DE"
```

**Liên kết file ghi âm** nằm ngay trong khung RECORDINGS ở đầu đề (nhấn để mở). Liên kết là đường dẫn tương đối `../Audio/…` nên cần giữ nguyên cấu trúc thư mục (không đổi tên/di chuyển file).

## Cấu trúc mỗi đề (20 câu, 10 điểm, 0,5 đ/câu)
- Part 1 (Q1–10): hoàn thành biểu mẫu, tối đa 3 từ và/hoặc số.
- Part 2 (Q11–15): chọn A/B/C; (Q16–20): nối với địa điểm A–E.
- Mỗi file chỉ phát một lần. Băng có lời dẫn (Test, Part, Questions…), 20 giây đọc đề, 30 giây kiểm tra đáp án; Part 2 có thêm 15 giây giữa Q15 và Q16. Mỗi file ≈ 2–2,5 phút (tổng 2 file ≈ 5 phút); thời gian thi ghi 10 phút gồm cả chép đáp án.

## Kiểm tra lại nội dung file nguồn (YouPass_Listening_Parts_1_2_Teacher_Edition_12_Tests.docx)
- Đã đối chiếu: 240 đáp án trong bảng đáp án ↔ 240 dấu [Q#: đáp án] trong transcript → khớp 100%; 12 test × (10 + 5 + 5) câu đầy đủ.
- **Link YouPass:** file nguồn ghi rõ câu hỏi/đáp án/transcript là bản **viết lại**, KHÔNG phải nguyên văn bài trên YouPass. Vì vậy audio trên YouPass sẽ không khớp đề này → đề dùng file ghi âm tạo từ chính transcript; link YouPass chỉ ghi trong file đáp án như nguồn tham chiếu. Danh sách mock test bạn gửi: https://youpass.vn/luyen-thi/ielts/listening?quiz_type=mocktest&status=unfinished&mock_source=youpass_collect (không truy cập được từ môi trường này nên chưa kiểm tra nội dung).
- **Lỗi đã sửa:** (1) đề Part 1 ghi "ONE WORD AND/OR A NUMBER" nhưng nhiều đáp án 2–3 từ (vd: Riverside Apartments, pop and jazz, Jim Granley) → đổi thành "NO MORE THAN THREE WORDS AND/OR A NUMBER"; (2) Test 7: các địa điểm Q16–20 trong băng không theo thứ tự câu hỏi → sắp lại theo thứ tự 16→20; (3) Test 8: 4/5 câu MC có đáp án B → đổi thứ tự phương án câu 12 và 14 (đáp án mới: B C A A B); (4) Test 5 câu 10: nhãn "Customer reads paper" → "How often the customer reads the paper".
- Part 1 trong nguồn khá dễ (đáp án nói ngay sau câu hỏi, ít nhiễu) – giữ nguyên nội dung theo yêu cầu dùng file nguồn.

## Lưu ý
- Giọng đọc là TTS tổng hợp (Piper, giọng Anh-Anh, 2 giọng cho hội thoại), không phải người thật. Nên nghe thử trước khi thi.
- Các file .docx chưa kiểm tra hiển thị bằng Word (môi trường không chạy được LibreOffice); PDF đã xem thử.
- Tạo lại: `_build_listening/` (`make_audio.py` → mp3; `build_listening.py` → docx/html; PDF bằng `_build/html2pdf.js`).
