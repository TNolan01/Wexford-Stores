WEXFORD STORES - installable search app

What is in this folder
  index.html, app.js, xlsx-reader.js   the app
  manifest.webmanifest, sw.js          make it installable and work offline
  icons/                               app icons

How to publish (GitHub Pages)
  1. Create a free account at github.com.
  2. New repository, name it wexford-stores, set to Public.
  3. Add file > Upload files. Drag in EVERYTHING from this folder (keep the icons folder). Commit.
  4. Settings > Pages > Branch: main, folder: / (root) > Save.
  5. After a minute or two the app is at https://YOURNAME.github.io/wexford-stores/
  6. Open that address in Chrome on Android, then menu (three dots) > Install app.

Netlify alternative: go to app.netlify.com/drop and drag this folder onto the page.

Using the app
  Tap Open file and choose an .xlsx or .csv. The first row is used as column headings,
  and the first column is the card title. The file stays on the phone and is remembered
  until you open another one. Old .xls files: in Excel use Save As and pick .xlsx.

Updating the app later
  Upload the changed files again and change CACHE_VERSION in sw.js (for example v1 to v2).
