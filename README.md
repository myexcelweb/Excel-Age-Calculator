# Excel Age Calculator (React + Vite)

Upload an Excel file, pick date columns, and download a new Excel with
Opening age (on a specific date), Disposal age (from date to date) and
Closing age (till today). Formats: Days / Months (2 dec) / Years (2 dec) / Y M D.

## Run locally
    npm install
    npm run dev

## Deploy on Netlify via Git push
1. Create a GitHub repo and push this folder:
       git init
       git add .
       git commit -m "Excel age calculator"
       git branch -M main
       git remote add origin https://github.com/<you>/<repo>.git
       git push -u origin main
2. Netlify -> Add new site -> Import from Git -> pick the repo.
3. Build settings are read from netlify.toml (build: `npm run build`, publish: `dist`).
4. Every later `git push` redeploys automatically.
