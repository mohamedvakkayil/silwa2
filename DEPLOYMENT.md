# Netlify Deployment Checklist

## Pre-Deployment

- [ ] Ensure `data.js` is up to date (run `python3 sync_excel.py`)
- [ ] Verify `e2.xlsx` is in the root directory
- [ ] Test locally using `python3 -m http.server 8000`
- [ ] Verify all features work:
  - [ ] Tree diagram loads
  - [ ] Search functionality
  - [ ] Detailed view opens
  - [ ] MDB View tabs work
  - [ ] Excel file loads for detailed views

## Files Required for Deployment

- [x] `index.html`
- [x] `script.js`
- [x] `styles.css`
- [x] `data.js` (generated from Excel)
- [x] `e2.xlsx` (source Excel file)
- [x] `netlify.toml` (configuration)
- [x] `.gitignore` (optional but recommended)

## Netlify Settings

1. **Build settings:**
   - Build command: (leave empty - static site)
   - Publish directory: `.` (root)

2. **Environment variables:**
   - None required

3. **Deploy settings:**
   - Auto-deploy: Enabled (recommended)
   - Branch: `main` or `master`

## Post-Deployment Verification

After deployment, verify:

1. **Homepage loads:**
   - [ ] Site loads at `https://your-site.netlify.app`
   - [ ] No console errors
   - [ ] Tree diagram displays

2. **Data loading:**
   - [ ] Data loads from `data.js`
   - [ ] Tree structure is correct
   - [ ] Statistics display correctly

3. **Interactive features:**
   - [ ] Click nodes to open detailed view
   - [ ] Search works
   - [ ] Filters work
   - [ ] MDB View tabs switch correctly
   - [ ] Collapsible rows work in MDB View

4. **Excel file access:**
   - [ ] Detailed views load Excel data
   - [ ] Special DB recalculation works
   - [ ] No CORS errors in console

## Troubleshooting

### Excel file not loading
- Check that `e2.xlsx` is committed to repository
- Verify file size is under Netlify's limits (100MB for free tier)
- Check browser console for errors

### Data not displaying
- Verify `data.js` is up to date
- Check browser console for JavaScript errors
- Ensure `data.js` is loaded before `script.js` in `index.html`

### CORS errors
- Netlify should handle CORS automatically via `netlify.toml`
- If issues persist, check Netlify function logs

### Build fails
- Check Netlify build logs
- Verify all required files are present
- Ensure no syntax errors in JavaScript

## Updating Data

When updating `e2.xlsx`:

1. Update Excel file locally
2. Run sync: `python3 sync_excel.py`
3. Commit changes:
   ```bash
   git add e2.xlsx data.js
   git commit -m "Update Excel data"
   git push
   ```
4. Netlify will auto-deploy

## Performance Tips

- `data.js` is cached by browsers (via `netlify.toml` headers)
- Excel file is only loaded when needed (detailed views)
- Consider compressing `e2.xlsx` if it's very large

