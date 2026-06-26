# 🌸 Smart Timetable System — Google Sheets + Apps Script

An automated, self-updating personal timetable system built entirely in Google Sheets and Apps Script. It live-syncs from a master institute schedule, intelligently resolves class conflicts, tracks real attendance, and presents everything in a clean pastel UI.

---

## ✨ Features

- **Live Sync** — pulls directly from the master institute timetable via Google Sheets API; no manual copy-pasting ever
- **Smart Conflict Resolution** — when two courses clash in the same slot, automatically decides which to attend using a skip-budget algorithm (protects whichever course has the least room left to skip)
- **80% Attendance Rule** — each course gets a skip budget of exactly 20% of its total sessions; the algorithm respects this across the whole term
- **Manual Attendance Override** — mark what you actually attended via a simple dropdown; future conflict decisions recalculate based on your real history, not just predictions
- **Exam & Quiz Tracking** — quizzes (Q1/Q2) and end-term exams (ET) are parsed from the master sheet, displayed in distinct gold/coral colors, and always marked compulsory
- **Today's Classes** — a dedicated tab that shows only today's schedule, pinned as the first tab, auto-refreshes every 12 hours
- **Weekly Summary** — attendance counts per subject per week with a pastel gradient heatmap showing busier weeks
- **Safety Net** — if a refresh finds zero sessions (broken link, renamed tab, etc.), it aborts before touching any existing data
- **Issues Log** — every refresh writes a timestamped report: what it found, any fuzzy-matched course codes, any warnings
- **Pretty Pastel UI** — consistent blush pink, lavender, and mint theme across all tabs using Quicksand font, rich text formatting, and color-coded attend/skip states

---

## 📋 Tabs

| Tab | Purpose |
|---|---|
| 📌 Today's Classes | At-a-glance view of today's schedule only |
| ✏️ Mark Attendance | Dropdown to mark Attended / Skipped for past classes |
| My Courses | Full term calendar grid — dates × slots |
| ⚠️ Conflicts | Skip budget overview + conflict-by-conflict decisions |
| 📊 Weekly Summary | Classes per subject per week with heatmap |
| Personal Days Off | Add dates you'll be absent entirely |
| 🩺 Issues Log | Auto-generated refresh diagnostics |

---

## 🛠️ Tools & Technologies

| Tool | Usage |
|---|---|
| Google Sheets | UI, data storage, tab rendering |
| Google Apps Script (JavaScript) | All logic — parsing, conflict resolution, styling |
| Spreadsheet API | Reading master sheet, writing all output tabs |
| Rich Text API | Strikethrough + color-coded text inside cells |
| Data Validation API | Attended / Skipped dropdowns in attendance tab |
| Conditional Formatting API | Pastel gradient heatmap on weekly summary |
| Time-driven Triggers | Auto-refresh every 12 hours |

---

## ⚙️ How It Works

```
Master Institute Sheet (live, read-only)
        ↓
  Apps Script reads & parses
        ↓
  ┌─────────────────────────────────┐
  │  For each date + slot:          │
  │  1. Check manual attendance     │
  │  2. Apply personal days off     │
  │  3. Detect clashes              │
  │  4. Run skip-budget algorithm   │
  │     → keep tightest budget      │
  └─────────────────────────────────┘
        ↓
  Rebuild all 7 tabs with styling
        ↓
  Log results to Issues Log
```

---

## 🚀 Setup (for your own use)

1. **Make a copy** of the Google Sheet template *(link your own here)*
2. Open **Extensions → Apps Script**
3. Paste the contents of `Code.gs`
4. Fill in the two settings at the top:
```javascript
const MASTER_URL = 'YOUR_MASTER_SHEET_URL_HERE';
const MASTER_TAB = 'YOUR_TAB_NAME_HERE';
```
5. Update `MY_COURSES` with your own course codes:
```javascript
const MY_COURSES = ['COURSE1', 'COURSE2', ...];
```
6. Update `EXAM_COURSE_MAP` with your quiz and exam codes:
```javascript
const EXAM_COURSE_MAP = {
  'Q1-COURSE1': { course: 'COURSE1', type: 'quiz' },
  'ET-COURSE1': { course: 'COURSE1', type: 'exam' },
  ...
};
```
7. Save → **Run → `refreshTimetable`** → approve permissions
8. Set up a trigger: **Triggers → Add Trigger → `refreshTimetable` → Every 12 hours**

---

## 📁 Repository Structure

```
├── Code.gs          # Full Apps Script source
└── README.md        # This file
```

---

## 💡 Key Design Decisions

- **Abort-before-overwrite** — a failed refresh never clears good data; it logs what went wrong and stops
- **Fuzzy course matching** — minor spelling/spacing differences in the master sheet are caught and logged as INFO rather than silently dropped
- **Exams are separate** — quizzes and end-terms are parsed and displayed but excluded from skip-budget logic and attendance tracking entirely
- **Manual marks persist** — the attendance tab is rebuilt on every refresh but your dropdown selections are re-read first and written back in, so nothing you marked is ever lost
- **Only past classes shown in attendance** — future sessions don't appear until their date arrives, keeping the tab short and actionable

---

## 📸 Screenshots

*(Add screenshots of each tab here with data blurred/cropped)*

---

## 👩‍💻 Author

Built as a personal productivity tool to manage a full-term MBA/PGP timetable with 9 concurrent courses, daily schedule updates, and 80% attendance requirements.
