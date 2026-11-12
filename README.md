# Amporia Works Billing

Counter billing for Amporia Works: load the price list, search components, build the bill and save the invoice as PDF or image.

Live: https://amporiaworks.github.io/billing/

## Files
- `index.html` – the page
- `css/style.css` – look and layout (brand colours, IBM Plex Mono)
- `js/app.js` – price-list reading, search, bill and invoice export
- `js/logo.js` – default invoice logo
- `assets/` – logos and app icons
- `manifest.webmanifest` – lets phones add it to the Home Screen like an app

## Notes
- The price list (Excel, CSV or PDF) is read in the browser and stored only on that device. Nothing is uploaded to GitHub.
- An internet connection is needed when the page opens (file-reading and PDF tools load from cdnjs, the font from Google Fonts).
