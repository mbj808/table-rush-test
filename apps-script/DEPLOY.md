# Deploy sheet script v9 (TEST backend only)
1. Sign in as mrjohnson808@gmail.com and open the TEST Apps Script project (web app AKfycbz2XyYe…/exec; sheet 1_x-dVyYlQ6SbKfaFvnZoY6KtXGpLuiGDBkdzwEFvNp8).
2. Replace all of Code.gs with apps-script/Code.v9.gs. Save.
3. Select `setupSessions` in the function menu and Run once (approve permissions if asked). It makes:
   - **Sessions** tab: one row per finished game: Start (HST) | Week of (Mon) | Date | Day | Class | Name | Game | Questions | Right | % correct | Minutes | Homework? | Start ms.
     Homework? = Yes outside 7:45 AM–2:05 PM Hawaii time on Mon–Fri, and all day Sat/Sun.
   - **Homework** tab (formulas, Homework?=Yes only): Week of | Class | Name | Nights played | Minutes | Questions | % correct | Games | Games played (e.g. "Multiplication ×5, Grammar ×2"). Filter by Class with a filter view.
4. Deploy > Manage deployments > edit (pencil) the existing web app > Version: New version > Deploy. Same deployment = same /exec URL.
5. Check: <exec URL>?action=version shows "version":9,"sessions":true.
Until then each device queues up to 200 game records and sends them once v9 is live.
To change school hours later: edit SCHOOL_START / SCHOOL_END near the v9 section (applies to new rows).
