# CCTV Daily Report (ICTU)

DOH Treatment and Rehabilitation Center Dulag. Live at https://trcdulag.github.io/cctv-daily-report/

- `index.html`: the whole app (no build step, works offline except Save to Google Sheet). The operator fills in the day's CCTV check and incidents, then taps **Save as PDF** (one A4 page) and/or **Save to Google Sheet**.
- `apps-script/Code.gs`: the backend in the "CCTV Daily Report Register (ICTU)" Google Sheet (salagfacility@gmail.com). After changing it, paste it into Extensions > Apps Script, then Deploy > Manage deployments > Edit > Version: New version > Deploy (keeps the same URL).

The web app URL is set as `DEFAULT_URL` in `index.html`.
