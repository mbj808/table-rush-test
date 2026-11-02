# Deploy sheet script v9 (TEST backend only)
1. Sign in as mrjohnson808@gmail.com, open the TEST Apps Script project (web app AKfycbz2XyYe…/exec, sheet 1_x-dVyYlQ6SbKfaFvnZoY6KtXGpLuiGDBkdzwEFvNp8).
2. Replace Code.gs with apps-script/Code.v9.gs. Save.
3. Run `setupSessions` once (makes the Sessions tab + Settings A14 "Teacher key"). Type a teacher key in Settings B14.
4. Deploy > Manage deployments > the existing web app > Edit (pencil) > Version: New version > Deploy. Keep the same deployment so the /exec URL stays the same.
5. Check: <exec URL>?action=version shows "version":9,"sessions":true.
Until then the site keeps session records queued on each device (up to 200) and sends them once v9 is live.
Report: https://mbj808.github.io/table-rush-test/homework.html (sample data: ?mock=1).
