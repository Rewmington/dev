# Final Exam Sync Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Update the final-exam portal so it adds a standalone `计算机组成原理` section, fully replaces the `数字逻辑与系统设计` section with the new choice-question bank, refreshes `毛泽东思想和中国特色社会主义理论体系概论` from the revised Word document, and shows subject-specific exam information from `期末.xlsx` inside each subject page.

**Architecture:** Keep the existing single-page data-driven renderer in `dev/final-exam/index.html`, but extend the `DATA` object with a new `computer` subject and a compact `examSchedule` map keyed by subject. Reuse the current chapter navigation and answer reveal system for ordinary choice questions, while adding a dedicated rendering branch for `计算机组成原理` so it shows stems only and no answer controls.

**Tech Stack:** HTML, vanilla JavaScript, `python-docx` for `.docx` extraction, and `pandas/openpyxl` for the Excel schedule.

---

### Task 1: Extract the source payloads

**Files:**
- Read: `D:\Desktop\期末题库\计算机\计算机.docx`
- Read: `D:\Desktop\期末题库\数字逻辑\选择.docx`
- Read: `D:\Desktop\期末题库\马院\毛泽东思想和中国特色社会主义理论体系概论复习材料.docx`
- Read: `D:\Desktop\期末题库\期末.xlsx`

- [ ] **Step 1: Confirm the four source shapes and counts**

```python
from docx import Document
import pandas as pd

computer_doc = Document(r'D:\Desktop\期末题库\计算机\计算机.docx')
dlogic_doc = Document(r'D:\Desktop\期末题库\数字逻辑\选择.docx')
mao_doc = Document(r'D:\Desktop\期末题库\马院\毛泽东思想和中国特色社会主义理论体系概论复习材料.docx')
schedule_df = pd.read_excel(r'D:\Desktop\期末题库\期末.xlsx')
```

- [ ] **Step 2: Validate the usable data that will feed the page**

Expected source payload:
- `计算机.docx` contributes 6 calculation/comprehensive prompts, with answers intentionally hidden in the UI for now.
- `选择.docx` contributes 15 single-choice items and 17 multiple-choice items.
- `毛概复习材料.docx` contributes the revised chapter text and answers already present in the review material.
- `期末.xlsx` contributes 6 schedule rows: 英语 / 毛概 / 习新思想 / Java高级编程 / 计算机组成原理 / 数字逻辑与系统设计.

- [ ] **Step 3: Build the subject-key schedule map that the page will render**

```javascript
const examSchedule = {
  eng: { label: '英语', date: '2026.07.06', time: '10:00 - 11:50', room: '2J 206' },
  mao: { label: '毛概', date: '2026.07.06', time: '14:00 - 15:50', room: '1J 402' },
  marx: { label: '习新思想', date: '2026.07.08', time: '14:00 - 15:50', room: '1J 402' },
  java: { label: 'Java高级编程', date: '2026.07.09', time: '08:00 - 09:50', room: '2J 503' },
  computer: { label: '计算机组成原理', date: '2026.07.09', time: '16:00 - 17:50', room: '2J 203' },
  dlogic: { label: '数字逻辑与系统设计', date: '2026.07.10', time: '10:00 - 11:50', room: '2J 204' }
};
```

- [ ] **Step 4: Record any source quirks before coding**

Examples to preserve while syncing:
- `毛概` uses Chinese punctuation and chapter headings from the revised Word doc.
- `计算机.docx` is answer-free in the web UI even though the document contains solution text.
- `数字逻辑` has separate single-choice and multiple-choice blocks only; the old mixed calculation content should be removed from the live page.

---

### Task 2: Add the `计算机组成原理` subject as a stems-only section

**Files:**
- Modify: `D:\Desktop\期末题库\dev\final-exam\index.html`

- [ ] **Step 1: Add the new subject to `DATA`**

```javascript
DATA.computer = {
  name: '计算机组成原理',
  icon: '🧮',
  color: '#22c55e',
  desc: '计算机组成原理 · 综合计算题',
  chapters: [
    {
      name: '综合计算题',
      questions: [
        { stem: '1.已知十进制X=-0.375，求8位定点小数原码、反码、补码。', options: [], answer: '' },
        { stem: '2.已知主频10MHz，每个机器周期包含5个时钟周期，整机平均指令速率0.5MIPS，求：机器周期、单指令平均执行时间、单指令占用机器周期数。', options: [], answer: '' }
      ]
    }
  ]
};
```

- [ ] **Step 2: Add a render branch that suppresses options and answer chips for this subject**

```javascript
if (currentSubject === 'computer') {
  // render only the stem and a small exam-info card per subject
  // do not render A/B/C/D options
  // do not render q-answer blocks
}
```

- [ ] **Step 3: Keep question order identical to the Word document**

Expected rendered order:
1. 定点小数原/反/补码
2. 机器周期与指令平均执行时间
3. 组相联 Cache 命中率与分组数
4. 双符号位补码加法与溢出判断
5. DRAM 刷新 / 乘法器 / 64KB 主存芯片组织
6. `MOV (R1),@(R2)+` 的分周期微操作 + 寻址方式

- [ ] **Step 4: Verify the section opens cleanly**

Run: open the page, select `计算机组成原理`.
Expected: the page shows question stems and the subject schedule card, but no answer toggles, no choice options, and no mixed old content.

- [ ] **Step 5: Commit**

```bash
git add D:\Desktop\期末题库\dev\final-exam\index.html
git commit -m "feat: add computer composition principles section"
```

---

### Task 3: Replace the digital-logic section with the new choice bank

**Files:**
- Modify: `D:\Desktop\期末题库\dev\final-exam\index.html`

- [ ] **Step 1: Replace `DATA.dlogic` with only the new single-choice and multiple-choice content**

```javascript
DATA.dlogic = {
  name: '数字逻辑与系统设计',
  icon: '🔢',
  color: '#f97316',
  desc: '数字逻辑与系统设计 · 选择题',
  chapters: [
    { name: '单项选择题', questions: [ /* 15 items from 选择.docx */ ] },
    { name: '多项选择题', questions: [ /* 17 items from 选择.docx */ ] }
  ]
};
```

- [ ] **Step 2: Preserve the existing single-choice / multi-choice answer interaction**

```javascript
// single-choice items keep the current A/B/C/D answer reveal
// multi-choice items keep the current multi-answer highlight behavior
```

- [ ] **Step 3: Remove the old calculation/analysis chapter from the live digital-logic subject**

```javascript
// delete the previous "分析计算题" chapter from DATA.dlogic
// keep the subject limited to the new question bank only
```

- [ ] **Step 4: Verify the counts before moving on**

Expected chapter counts:
- `单项选择题` = 15
- `多项选择题` = 17

- [ ] **Step 5: Commit**

```bash
git add D:\Desktop\期末题库\dev\final-exam\index.html
git commit -m "feat: replace digital logic question bank"
```

---

### Task 4: Refresh the Mao Zedong Theory subject from the revised Word file

**Files:**
- Modify: `D:\Desktop\期末题库\dev\final-exam\index.html`

- [ ] **Step 1: Update the Mao subject text and answers to match the revised Word document**

```javascript
DATA.mao = {
  name: '毛泽东思想和中国特色社会主义理论体系概论',
  icon: '📖',
  color: '#f97316',
  desc: '毛泽东思想和中国特色社会主义理论体系概论 · 复习材料',
  chapters: [
    // keep the existing chapter partitions, but replace the text/answers with the revised docx content
  ]
};
```

- [ ] **Step 2: Keep the chapter segmentation that the page already uses**

```javascript
// 导论 / 第二章 / 第三章 / 第四章 / 第五章 / 第七章 / 第八章 should remain separate chapters
// only the content and answer text should be refreshed
```

- [ ] **Step 3: Validate the refreshed wording against the Word source**

Expected: the updated chapter copy and answer keys match the revised `毛概复习材料.docx` exactly, including punctuation and multi-select answer strings.

- [ ] **Step 4: Commit**

```bash
git add D:\Desktop\期末题库\dev\final-exam\index.html
git commit -m "feat: refresh mao review content"
```

---

### Task 5: Add per-subject exam info cards from `期末.xlsx`

**Files:**
- Modify: `D:\Desktop\期末题库\dev\final-exam\index.html`

- [ ] **Step 1: Add a reusable exam-info renderer**

```javascript
function renderExamInfo(subjectKey) {
  const info = examSchedule[subjectKey];
  if (!info) return '';

  return `
    <div class="ep-panel">
      <div class="ep-header">
        <h3>📅 本学科考试信息</h3>
      </div>
      <div class="ep-body open">
        <table>
          <tr><th>考试科目</th><td>${info.label}</td></tr>
          <tr><th>日期</th><td>${info.date}</td></tr>
          <tr><th>时间</th><td>${info.time}</td></tr>
          <tr><th>考试地点</th><td>${info.room}</td></tr>
        </table>
      </div>
    </div>`;
}
```

- [ ] **Step 2: Render the card when a subject opens**

```javascript
html += renderExamInfo(currentSubject);
```

- [ ] **Step 3: Add a fallback for subjects that do not have a schedule row**

```javascript
if (!examSchedule[currentSubject]) {
  html += '<div class="q-card"><div style="color:var(--text2)">暂无对应考试安排。</div></div>';
}
```

- [ ] **Step 4: Verify every mapped subject shows the right row from the spreadsheet**

Expected mapping:
- `eng` → 英语 / 2026.07.06 / 10:00 - 11:50 / 2J 206
- `mao` → 毛概 / 2026.07.06 / 14:00 - 15:50 / 1J 402
- `marx` → 习新思想 / 2026.07.08 / 14:00 - 15:50 / 1J 402
- `java` → Java高级编程 / 2026.07.09 / 08:00 - 09:50 / 2J 503
- `computer` → 计算机组成原理 / 2026.07.09 / 16:00 - 17:50 / 2J 203
- `dlogic` → 数字逻辑与系统设计 / 2026.07.10 / 10:00 - 11:50 / 2J 204

- [ ] **Step 5: Commit**

```bash
git add D:\Desktop\期末题库\dev\final-exam\index.html
git commit -m "feat: add subject exam schedule cards"
```

---

### Task 6: Smoke test and final polish

**Files:**
- Modify if needed: `D:\Desktop\期末题库\dev\final-exam\index.html`
- Test in browser

- [ ] **Step 1: Open the portal and switch through all subjects**

Run: open `dev/index.html`, then enter the final-exam portal and click through `Java / 毛概 / 习新思想 / 数字逻辑与系统设计 / 计算机组成原理 / 大学英语4`.
Expected: no console errors and every subject shows its own exam-info card.

- [ ] **Step 2: Inspect `计算机组成原理` for stems-only rendering**

Run: click the new subject.
Expected: no answer chips, no option buttons, and no legacy calculation section from `数字逻辑`.

- [ ] **Step 3: Inspect the new `数字逻辑` banks**

Run: open both choice chapters.
Expected: exactly 15 single-choice items and 17 multiple-choice items, with answer reveal still working.

- [ ] **Step 4: Inspect the Mao refresh**

Run: open `毛泽东思想和中国特色社会主义理论体系概论`.
Expected: updated text/answer content from the revised Word source.

- [ ] **Step 5: Final commit after all checks pass**

```bash
git add D:\Desktop\期末题库\dev\final-exam\index.html
# add any helper data file only if one was created during implementation
git commit -m "feat: sync final exam portal with latest sources"
```
