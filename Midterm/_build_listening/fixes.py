# -*- coding: utf-8 -*-
"""Chỉnh sửa/kiểm tra nội dung so với file nguồn YouPass_Listening_Parts_1_2_Teacher_Edition_12_Tests.docx"""
P1_DIRECTIONS="Complete the notes/form below. Write NO MORE THAN THREE WORDS AND/OR A NUMBER for each answer."
LABEL_FIX={(5,10):"How often the customer reads the paper"}
# Test 7: các câu nối 16-20 trong script nghe không theo thứ tự câu hỏi -> sắp lại theo thứ tự 16..20
SCRIPT2_FIX={7:["Our recycling centre has been redesigned so that we can accept a wider range of household waste. [Q11: A] We take small electrical items and household batteries, but commercial refrigerators must be handled by a specialist contractor. [Q12: B] If you bring paint, please speak to a member of staff before unloading it. [Q13: B] The reuse shop is free to enter; delivery and repairs cost extra. [Q14: C] The quietest time is usually early on Tuesday. [Q15: A]",
 "Batteries belong in Zone 2. [Q16: B] Small electrical appliances go to Zone 4. [Q17: D] Glass goes in Zone 1. [Q18: A] Garden waste is collected in Zone 5. [Q19: E] Paint is in Zone 3. [Q20: C]"]}
# Test 8: 4/5 câu MC có đáp án B -> đổi thứ tự phương án câu 12 và 14 để cân bằng
MC_REORDER={(8,12):['C','A','B'],(8,14):['B','C','A']}   # new A,B,C = old options in this order
# Thêm đánh vần họ/tên (đề nguồn không đánh vần nhiều tên riêng – không công bằng cho phần nghe chính tả)
SCRIPT1_REPL={3:[("Guest: Bennett. [Q1: Bennett]","Guest: Bennett — B-E-N-N-E-T-T. [Q1: Bennett]")],
 6:[("The contact person for new members is Jim Granley. [Q1: Jim Granley]","The contact person for new members is Jim Granley — G-R-A-N-L-E-Y. [Q1: Jim Granley]")],
 10:[("Employee: Patel. [Q1: Patel]","Employee: Patel — P-A-T-E-L. [Q1: Patel]")],
 11:[("Competitor: Dawson. [Q1: Dawson]","Competitor: Dawson — D-A-W-S-O-N. [Q1: Dawson]")],
 12:[("Runner: Lewis. [Q1: Lewis]","Runner: Lewis — L-E-W-I-S. [Q1: Lewis]")]}
