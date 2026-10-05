# GPSC Portal

## Who can do what
| | Parents / students | Library | Faculty | Admin (one account) |
|---|---|---|---|---|
| See own attendance + test results (home page, no login) | yes | | | |
| Add library attendance (in / out / book) | | yes | | yes |
| Add lecture attendance + test marks (own batches) | | | yes | yes |
| Edit / delete saved attendance, library entries, tests | | | | yes |
| Create batches (GPSC 1&2 2026, TALATI ...) | | | | yes |
| Add / edit / deactivate students (one by one or bulk paste) | | | | yes |
| Create staff, assign role (faculty / library) and batches | | | | yes |

Faculty and library staff only see the batches the admin assigned to them. Once a faculty/library
member saves an entry for a date, it is locked for them; only the admin can change or delete it.

## Run on your PC (Windows)
Double-click `run.bat`, open http://localhost:8000, log in as `admin` (password asked on first start).
Then: "બેચ" tab -> create batch -> "વિદ્યાર્થીઓ" tab -> add students -> "સ્ટાફ" tab -> create faculty / library logins.

## Put it online (free)
1. neon.com -> new project -> copy the connection string (postgresql://...)
2. Upload this folder (without run.bat) to a private GitHub repo
3. render.com -> New -> Blueprint -> pick repo -> enter ADMIN_PASSWORD and DATABASE_URL
4. Open the .onrender.com link on a phone -> Install app
Optional: UptimeRobot pinging https://YOUR-APP.onrender.com/api/health every 5 min avoids the free-plan sleep.

## Notes
- Upgrading from the previous version keeps your old data in tables named legacy_* and starts a fresh
  default batch ("GPSC બેચ", the 51 students); existing staff logins keep working.
- Forgot the admin password? Set RESET_ADMIN_PASSWORD=1 and ADMIN_PASSWORD=<new> once, start, then remove them.
- The public page never shows mobile numbers or parent details; only admin sees them.
