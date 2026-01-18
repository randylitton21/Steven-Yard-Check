# Yard Check App

This is a static web app (HTML/CSS/JS) designed to run on GitHub Pages or any
static host.

## Run Locally
- Open `index.html` in a browser.

## GitHub Pages Deployment
1. Create a new GitHub repository (public).
2. Initialize and push this folder:
   - `git init`
   - `git add .`
   - `git commit -m "Initial commit"`
   - `git branch -M main`
   - `git remote add origin https://github.com/<user>/<repo>.git`
   - `git push -u origin main`
3. In GitHub: Settings → Pages → Source: `main` branch, `/ (root)`.
4. Your site will be live at:
   - `https://<user>.github.io/<repo>/`

## Notes
- Keep `index.html`, `styles.css`, `app.js`, and the image assets in the root.
- If you rename files, update the references in `index.html`.
