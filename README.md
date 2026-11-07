# RTU Grade Portal (Google Apps Script)

One web app with two pages:

| Page | URL | Who |
|---|---|---|
| **Student Portal** | `…/exec` | Students signed in with `studentnumber@rtu.edu.ph` |
| **Instructor Page** | `…/exec?page=instructor` | Only you (the owner) and co-instructors you list |

One record book holds **all your subjects/sections, one per tab**. Students see every released subject they are enrolled in.

## Files

| File | Purpose |
|---|---|
| `Code.gs` | Server code: reads every tab, checks the student, sends emails, keeps logs |
| `Instructor.html` | Record Book (subjects + release switches), Class Records, Message & Email, Portal Settings |
| `Student.html` | Email check → student number + name → Submit → subjects, grades, consultation booking |
| `Styles.html` | Shared look: Beige / Royal Blue / Burgundy |
| `appsscript.json` | Manifest: time zone, permissions, web-app settings |

## Setup (about 10 minutes)

1. Sign in with your **@rtu.edu.ph** Google account (the one that owns the record book) and go to <https://script.google.com>. Click **New project** and name it *Grade Portal*.
2. Paste **Code.gs** into `Code.gs`.
3. Click **＋ › HTML** to add **Instructor**, **Student**, and **Styles**. Leave off the `.html` because Apps Script adds it. Paste in each file's contents.
4. Under **Project Settings ⚙**, check **Show "appsscript.json" manifest file in editor**. Replace its contents with **appsscript.json**.
5. Click **Deploy › New deployment › Web app** and set:
   - *Execute as:* **Me**
   - *Who has access:* **Anyone within Rizal Technological University**
   - Click **Deploy** and **authorize** the permissions.
6. Open the web-app URL with `?page=instructor` at the end:
   1. **Record Book:** paste your record book link and click **Link**. Every subject tab is listed.
   2. Turn **Released** on only for subjects whose grades are final, then click **Save**.
   3. **Class Records:** pick a subject and click a student to see exactly what they'll see.
   4. **Message & Email:** edit the message, then **Send test to me**.
   5. **Portal Settings:** check your consultation booking link.
   6. Share the **Student portal** link with all your classes.

> After you edit the code, use **Deploy › Manage deployments › ✏️ › Version: New version** to keep the same URL.

## How your record book is read

The reader is built for your class-record template, and it works on any tab laid out the same way:

| What | Where (your template) |
|---|---|
| Course, Section, Instructor | `COURSE:`, `YEAR AND SECTION:`, `INSTRUCTOR:` above the table |
| Headers | Row 10 (category), row 11 (date), row 12 (item); merged cells are handled |
| Highest possible scores | The **SCORES** row (row 13), so scores show as `38 / 50` |
| Students | Every row whose STUDENT NUMBER looks like `2026-101234` |
| **Summary of Grades** | Columns between LEARNERS' NAMES and Initial Grade (Long Exams, Recitation, Written Works, Attendance: totals and weighted scores) |
| **Breakdown of Grades** | Columns after the grades (Quizzes, Exercises, Problem Sets/Assignments, Attendance by date) |
| Remarks | Computed from the Semestral Grade (1.00–3.00 Passed, 5.00 Failed, INC, DRP). If **Adjusted Grade** is filled in, it is shown and remarks use it instead |

- Columns with **no scores yet** (e.g. E5–E11, PS2–PS7) are hidden and appear once you enter scores.
- Attendance `1` / `0` shows as **Present** / **Absent**.
- **New tabs** (e.g. a new section) appear automatically. They start **unreleased**.
- Students see sheet edits within about **2 minutes** because of a short cache. The instructor page always reads fresh data.

## Student check

1. The page detects the signed-in Google account. Anything that isn't `@rtu.edu.ph` is blocked.
2. The student enters their **Student Number**, **Last / First / Middle Name**, and clicks **Submit**.
3. **Only the student number decides access.** It must match the part before `@` in their email. The name is saved in the log but never checked.
4. They see each released subject they're enrolled in, with Semestral Grade, Initial Grade, Remarks, Summary, Breakdown, and your message.

## Consultation (Failed remarks only)

- The **Request Consultation** button appears only for a subject where the student's remarks are **Failed**.
- It opens your **Google Calendar appointment page** (`https://calendar.app.google/7zMwLhfQeC39F5Ea6`). Only the open slots in your consultation schedule can be booked. To change your hours, edit the appointment schedule in Google Calendar; the portal doesn't need to change.
- The student may add a short concern first. It's saved in the **Consultation Requests** tab, and you can optionally receive an email (Portal Settings).

## Logs added to your record book

These tabs are created automatically and ignored as subjects: **Portal Log**, **Consultation Requests**, and **Email Log**.

## Short link for students: `mrvitalez.github.io/Grade-Distribution`

`docs/index.html` is a small GitHub Pages page that forwards students to the Apps Script portal. It holds no student data.

1. **Set the portal URL:** on GitHub, open `docs/index.html`, click ✏️, and replace `PASTE_YOUR_WEB_APP_URL_HERE` with your web-app URL (the one ending in `/exec`). Commit.
2. **Turn on Pages:** go to **Settings › Pages**, set *Source* to **Deploy from a branch**, and choose Branch **main** and folder **/docs**. Click **Save**.
3. After a minute or two, share **https://mrvitalez.github.io/Grade-Distribution/** with your classes.

If you create a **new deployment**, the `/exec` URL changes, so update step 1 again. Choosing a *new version* of the existing deployment keeps the same URL.

## Notes

- **Email quota:** about 1,500 emails a day on Workspace accounts. Emails are sent one subject at a time.
- **Instructor preview:** opening the Student Portal yourself lets you enter any student number, including for unreleased subjects.
- **Embedding in Google Sites:** use *Insert › Embed › By URL* with the web-app URL.
- **Palette:** the hex values in `Styles.html` and `Code.gs` are screen estimates of Pantone P 17-1 C, P 102-8 C, and P 50-16 C. Adjust them if you have official values.
