# RW B2.2 – MIDTERM TEST – READING

6 mã đề (RW22-R-01 … RW22-R-06), mỗi đề **25 câu / 2 bài đọc / 40 phút / 10 điểm (0,4 đ/câu)**: Passage 1 = câu 1–12 (gốc Q14–25), Passage 2 = câu 13–25 (gốc Q27–39). Đã **bỏ câu cuối của mỗi passage** (gốc Q26 và Q40). Bố cục theo bộ Grammar/Listening (header Học viện, ô MÃ ĐỀ, phiếu trả lời cuối đề), có bản .docx và .pdf.

```
RW B2.2-MIDTERM TEST - READING/
├── De_thi_va_Phieu_tra_loi/   RW22-R-01 … 06 - DE THI (.docx + .pdf)
└── Dap_an/                    RW22-R-01 … 06 - DAP AN (.docx + .pdf, có giải thích) + 00 - BANG DAP AN NHANH 6 MA DE
```

| Mã đề | Passage 1 (Q14–25 gốc) | Passage 2 (Q27–39 gốc) |
|---|---|---|
| 01 | Computer Provides More Questions Than Answers | Mystery in Easter Island! |
| 02 | Corporate Social Responsibility | The Exploration of Mars |
| 03 | Ancient Storytelling | Paper or Computer? |
| 04 | Leaf-Cutting Ants and Fungus | Assessing the risk |
| 05 | Left-handed or Right-handed | The Power of Nothing |
| 06 | Western Immigration of Canada | Beyond the Blue Line |

## Nguồn và cách kiểm tra
- Nguồn: Google Drive của giảng viên, thư mục **IELTS Reading Actual Tests 1** (các file "NN Passage 2/3 - … .pdf", có đủ bài đọc, câu hỏi, đáp án). Đã dùng Test 16, 17, 18, 19 (IELTS 19), 20, 22. **Test 21** bị loại vì bài đọc bị lỗi/lẫn câu và đáp án không đáng tin.
- Các trang YouPass/ieltsreading.info/study4 **không truy cập được** từ môi trường làm việc; file "YouPass_Collect_…_Teacher_Key" chỉ có đáp án (không có bài đọc/câu hỏi) nên không dùng để soạn đề.
- Từng câu đã được đối chiếu với bài đọc. **Đáp án nguồn bị sửa (3 câu):** Test 17 Q36 B→A (Mars/SOLID); Test 16 Q31 NOT GIVEN→TRUE (Easter Island: cư dân đầu tiên là người Polynesia); Test 22 Q39 FALSE→NOT GIVEN (Lapita/Fiji). Cũng chỉnh câu 36 Test 22 thành "against the prevailing wind". Các câu còn mơ hồ nhưng giữ đáp án nguồn có ghi chú trong file đáp án: Test 16 Q16–17, Test 18 Q37, Test 20 Q20 và Q31.
- Văn bản trích từ PDF có vài lỗi (thiếu/lặp chữ, câu chen ngang…) đã được chỉnh nhẹ (vd: "war threatened", "He found it was not inert wax", bỏ đoạn lặp ở Test 20 Passage 3 đoạn E, bỏ đoạn "Sellen and Harper write:" bị cụt ở Test 18, sửa chính tả…). Nên đọc lướt bài đọc trước khi in.
- Câu hỏi được giữ nguyên nội dung gốc; một số chỗ chỉnh lỗi ngữ pháp nhẹ trong câu hỏi, và viết lại hướng dẫn cho rõ.
- Header ghi **RW B2.2 (Reading & Writing B2.2)**, 40 phút – giảng viên xác nhận/sửa trong `_build_reading/build_reading.py`.

Tạo lại: `cd Midterm/_build_reading && python3 build_reading.py "<thư mục>"` (HTML_OUT=<dir> cho PDF, rồi `node ../_build/html2pdf.js`).
