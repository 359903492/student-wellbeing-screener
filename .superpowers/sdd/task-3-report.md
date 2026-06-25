# Task 3 Report: JavaScript Interaction Logic for student.html

## Status: COMPLETE

## What was implemented

All JavaScript code from the implementation plan (task-3-brief.md) was transcribed into the `<script>` tag of `student.html`. The implementation covers four areas:

### 1. Question Data and Configuration
- `API_URL` placeholder set to `'YOUR_GOOGLE_APPS_SCRIPT_URL'`
- `QUESTIONS` object with PHQ-9 (9 questions) and GAD-7 (7 questions) using standard Chinese wording verbatim from the plan
- `OPTIONS` array: 完全不会 (0), 好几天 (1), 一半以上天数 (2), 几乎每天 (3)
- `ALL_QUESTIONS` merged array with `group` property for scoring
- `answers` object for tracking responses, `currentQuestion` for state

### 2. Question Rendering and Navigation
- `showScreen(id)` toggles `.active` class on screen divs
- `renderQuestion(index)` renders one question at a time with auto-advance on option click (300ms delay)
- Section headers ("PHQ-9 情绪状态", "GAD-7 焦虑状态") shown on first question and when section changes
- Progress bar updates in real-time
- Submit button (`#btn-submit`) revealed only after all 16 questions are answered
- Completion message shown when all questions done

### 3. Scoring and Level Classification
- `collectAnswers()` extracts PHQ-9 (9 scores) and GAD-7 (7 scores) arrays
- PHQ-9 severity levels: 0-4 normal, 5-9 mild, 10-14 moderate, 15-19 moderately severe, 20-27 severe
- GAD-7 severity levels: 0-4 normal, 5-9 mild, 10-14 moderate, 15-21 severe
- Level objects: `{ level: 'good'|'mild'|'moderate'|'severe', label: '...' }`

### 4. API Submission and Result Display
- `submitAnswers()` POSTs to API_URL with `{ action: "submit", studentId, phq9, gad7 }`
- On success: calls `showResult()` with server-returned totals
- `showResult()` renders two "temperature meters" with level badges
- Meter fill colors: good=green, mild=yellow, moderate=orange, severe=red
- Severe-level results show extra warning card with resource suggestions
- Error handling: alerts for submission failure and network errors with button re-enable

### 5. Event Initialization
- Student ID input: enables start button when non-empty
- Start button click: validates student ID, navigates to questions screen
- Submit button click: triggers `submitAnswers()`
- Retry button: reloads page (inline onclick, pre-existing in HTML)

## CSS fix applied

Changed `.meter-fill.low` to `.meter-fill.good` in the `<style>` block to match the `level` values returned by `getPHQ9Level()` and `getGAD7Level()` (which use `'good'` not `'low'`). Without this fix, the "normal range" meter bar would render without color.

## User flow verification

Full trace-through of the complete user flow passed:
- Consent screen with disabled start button until student ID entered
- 16 questions rendered one at a time with auto-advance
- Section headers shown only when section changes (PHQ-9 at Q1, GAD-7 at Q10)
- Submit button revealed only after all questions answered
- Results display with temperature meters, level badges, and disclaimer
- Severe results show additional warning with help resources
- All scoring thresholds verified against requirements

## Fix applied (2026-06-25): PHQ-9 Q9 self-harm warning always shows

### Issue
PHQ-9 question 9 (self-harm: "有不如死掉或用某种方式伤害自己的念头") score >= 1 must show an extra warning card in the results screen, regardless of total severity level. Previously the warning only showed when total level was 'severe'.

### Changes made (commit ead0499)

1. **`showResult` signature** -- added `phq9Q9Score` as the fifth parameter:
   ```
   function showResult(phq9Total, phq9Level, gad7Total, gad7Level, phq9Q9Score)
   ```

2. **Q9 safety check** -- inserted after the two temperature meter cards and before the existing severe-level check:
   ```javascript
   if (phq9Q9Score >= 1) {
     html += '<div class="card" style="background:#FFF5F5;border:2px solid #FFCDD2;text-align:center;">' +
       '<p style="font-size:15px;color:#C62828;">💛 你的回答中有一项引起了我们的关注。</p>' +
       '<p style="font-size:13px;color:#C62828;margin-top:8px;">请务必找班主任或学校心理老师聊一聊。你很重要，有人愿意帮你。</p>' +
       '</div>';
   }
   ```
   This card uses a distinct red-tinted style (`#FFF5F5` background, `#FFCDD2` border, `#C62828` text) to differentiate it from the generic severe-level warning (`#FFF3E0` background).

3. **Submit callback** -- passes `answers.phq9[8] || 0` as the fifth argument to `showResult()`. The `answers` variable is already in the closure scope from the enclosing `submitAnswers()` function, so no additional data plumbing is needed.

### Verification
- JavaScript syntax validated with `new Function()` -- no parse errors
- Both warning cards are independent: the Q9 safety card triggers on `phq9Q9Score >= 1`; the severe-level card triggers on `phq9Level.level === 'severe' || gad7Level.level === 'severe'`. They can appear together or separately.
