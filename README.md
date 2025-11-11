# Silwa Tower Estimation - Technoplus

A responsive web application for visualizing and managing electrical distribution board estimates for the Silwa Tower project.

## Features

- Interactive D3.js tree diagram showing electrical distribution hierarchy
- Detailed view for each component with metadata and Excel sheet data
- MDB View with tabular data for MDB1, MDB2, MDB3, and MDB4
- Collapsible nodes for SMDB, ESMDB, and MCC
- Special DB recalculation based on parent SMDB and copy counts
- Real-time search and filtering capabilities

## Deployment on Netlify

### Prerequisites

1. A Netlify account (free tier works fine)
2. Your project files committed to a Git repository (GitHub, GitLab, or Bitbucket)

### Deployment Steps

1. **Prepare your repository:**
   - Ensure `data.js` is up to date (run `python3 sync_excel.py` locally if needed)
   - Commit all files including:
     - `index.html`
     - `script.js`
     - `styles.css`
     - `data.js`
     - `e2.xlsx`
     - `netlify.toml`

2. **Deploy to Netlify:**
   - Go to [Netlify](https://app.netlify.com)
   - Click "Add new site" → "Import an existing project"
   - Connect your Git repository
   - Netlify will auto-detect settings from `netlify.toml`
   - Click "Deploy site"

3. **Verify deployment:**
   - Once deployed, your site will be live at `https://your-site-name.netlify.app`
   - Test the site to ensure:
     - Data loads correctly
     - Excel file (`e2.xlsx`) loads for detailed views
     - All interactive features work

### Updating Data

When you update `e2.xlsx`:

1. **Local update:**
   ```bash
   python3 sync_excel.py
   ```

2. **Commit and push:**
   ```bash
   git add data.js e2.xlsx
   git commit -m "Update Excel data"
   git push
   ```

3. **Netlify will automatically redeploy** when it detects the push

### Manual Sync Script

If you need to sync Excel data manually:

```bash
# Run the sync script
python3 sync_excel.py

# Or use the shell script
./sync.sh
```

## Local Development

### Requirements

- Python 3.x (for Excel sync scripts)
- A local HTTP server

### Running Locally

1. **Sync Excel data:**
   ```bash
   python3 sync_excel.py
   ```

2. **Start a local server:**
   ```bash
   python3 -m http.server 8000
   ```

3. **Open in browser:**
   ```
   http://localhost:8000
   ```

### Auto-sync (Development)

For automatic syncing during development:

```bash
python3 watch_excel.py
```

This will watch `e2.xlsx` for changes and automatically regenerate `data.js`.

## File Structure

```
.
├── index.html          # Main HTML file
├── script.js           # Main JavaScript logic
├── styles.css          # Stylesheet
├── data.js             # Generated data from Excel (auto-generated)
├── data.json           # JSON version of data (fallback)
├── e2.xlsx             # Source Excel file
├── netlify.toml        # Netlify configuration
├── sync_excel.py       # Script to sync Excel to data.js
├── watch_excel.py      # Auto-watch script for development
└── README.md           # This file
```

## Browser Support

- Chrome/Edge (latest)
- Firefox (latest)
- Safari (latest)
- Mobile browsers (iOS Safari, Chrome Mobile)

## Notes

- The site uses `data.js` as the primary data source (embedded JavaScript)
- Excel file (`e2.xlsx`) is loaded dynamically for detailed views
- All paths are relative, making it compatible with Netlify's static hosting
- No build process required - pure static HTML/CSS/JavaScript

## Support

For issues or questions, contact Technoplus development team.

