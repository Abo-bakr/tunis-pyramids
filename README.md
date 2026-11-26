# Tunis Pyramids P&L dashboard

Live page: https://abo-bakr.github.io/tunis-pyramids/

`index.html` is rebuilt automatically. Every 30 minutes the GitHub Actions workflow
**Update dashboard from Google Drive** reads the hotel system's monthly
"Profit and Loss" exports (.xls or .xlsx) from the Google Drive upload folder,
recalculates every figure, and publishes the page if anything changed.

## Monthly routine

1. Export the month's Profit and Loss from the hotel system.
2. Upload the file to the Google Drive folder **Tunis Pyramids - P&L uploads**.
3. Within about 30 minutes the dashboard shows the new month. To update at once,
   open **Actions → Update dashboard from Google Drive → Run workflow**.

Only files that contain the system's "Profit and Loss" report are used; other files
in the folder are ignored. If two files cover the same month, the newest one wins.

## Settings (repository secrets)

- `GDRIVE_SERVICE_ACCOUNT_KEY`: the Google service account's JSON key.
- `GDRIVE_FOLDER_ID`: the id of the Drive upload folder.

## Files

- `build.js`: reads the exports and builds `index.html`.
- `page-template.html`, `shell.html`: the dashboard page the script fills in.
- `.github/workflows/update.yml`: the schedule.

Run locally: `npm install` then `node build.js --local <folder-with-exports>`.
