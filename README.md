# DataGen Statistics - Office Add-in

A modern, cross-platform replacement for the DataGen Excel VBA macros. This Office Add-in provides real-time descriptive statistics and data generation capabilities that work on Excel for Web, Desktop (Windows/Mac), iOS, and Android.

## Features

- **Descriptive Statistics Panel**: Real-time calculations for up to 5 variables
  - N, Mean, Median, Mode
  - Standard Deviation, Variance
  - Min, Max, Range
  - Sum, Sum of Squares

- **Inferential Statistics**:
  - Correlation (r) and r²
  - Linear regression (slope, intercept)
  - Independent samples t-test
  - Paired samples t-test

- **Data Generation**:
  - Normal distribution (Box-Muller transform)
  - Uniform distribution
  - Sequences
  - Configurable decimal places

- **Real-time Updates**: Statistics update automatically when data changes

## Platform Support

| Platform | Status |
|----------|--------|
| Excel for Web | ✅ Supported |
| Excel Desktop (Windows) | ✅ Supported |
| Excel Desktop (Mac) | ✅ Supported |
| Excel for iOS | ✅ Supported |
| Excel for Android | ✅ Supported |

## Installation

### Option 1: Sideload for Development/Testing

1. Clone this repository or download the files
2. Host the files on a web server (or use GitHub Pages)
3. Update `manifest.xml` with your hosting URL
4. Sideload the manifest in Excel:
   - **Excel Desktop**: File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs
   - **Excel Web**: Insert → Office Add-ins → Upload My Add-in

### Option 2: GitHub Pages Hosting

1. Fork this repository
2. Enable GitHub Pages in repository settings (use `main` branch)
3. Update `manifest.xml`:
   - Replace all instances of `YOUR_GITHUB_USERNAME` with your actual GitHub username
4. Wait for GitHub Pages to deploy (usually 1-2 minutes)
5. Sideload the manifest.xml file in Excel

## Setup Instructions

### Step 1: Update the Manifest

Edit `manifest.xml` and replace `YOUR_GITHUB_USERNAME` with your GitHub username:

```xml
<!-- Change this: -->
<SourceLocation DefaultValue="https://YOUR_GITHUB_USERNAME.github.io/datagen-addin/taskpane.html"/>

<!-- To this (example): -->
<SourceLocation DefaultValue="https://johndoe.github.io/datagen-addin/taskpane.html"/>
```

### Step 2: Create Icon Files

The add-in needs PNG icon files. Create these in the `assets` folder:
- `icon-16.png` (16x16 pixels)
- `icon-32.png` (32x32 pixels)
- `icon-64.png` (64x64 pixels)
- `icon-80.png` (80x80 pixels)

You can use the included `icon.svg` as a template.

### Step 3: Deploy to GitHub Pages

1. Push all files to your GitHub repository
2. Go to Settings → Pages
3. Select "Deploy from a branch"
4. Choose `main` branch and `/ (root)` folder
5. Click Save

### Step 4: Sideload the Add-in

#### Excel for Web:
1. Go to excel.office.com and open a workbook
2. Click Insert → Office Add-ins
3. Click "Upload My Add-in"
4. Select your `manifest.xml` file
5. Click "DataGen" in the ribbon to open the panel

#### Excel Desktop (Windows):
1. File → Options → Trust Center → Trust Center Settings
2. Click "Trusted Add-in Catalogs"
3. Add your GitHub Pages URL as a catalog
4. Restart Excel
5. Insert → My Add-ins → DataGen Statistics

#### Excel Desktop (Mac):
1. Insert → Add-ins → My Add-ins
2. Click "Upload My Add-in"
3. Select your `manifest.xml` file

## Data Structure

The add-in expects data in this format (matching the original DataGen VBA):

| Row | A | B | C | D | E | F |
|-----|---|---|---|---|---|---|
| 3 | | Variable 1 | Variable 2 | Variable 3 | Variable 4 | Variable 5 |
| 4 | 1 | (data) | (data) | (data) | (data) | (data) |
| 5 | 2 | (data) | (data) | (data) | (data) | (data) |
| ... | ... | ... | ... | ... | ... | ... |
| 203 | 200 | (data) | (data) | (data) | (data) | (data) |

- **Row 3**: Variable names (headers)
- **Rows 4-203**: Data (up to 200 rows per variable)
- **Columns B-F**: Five variables

## File Structure

```
datagen-addin/
├── manifest.xml        # Add-in configuration (edit with your GitHub username)
├── taskpane.html       # Main UI panel
├── taskpane.js         # JavaScript logic and Office.js integration
├── taskpane.css        # Styles
├── functions.html      # Required by manifest (minimal)
├── assets/
│   ├── icon.svg        # Source icon (convert to PNG)
│   ├── icon-16.png     # Ribbon icon (16x16)
│   ├── icon-32.png     # Ribbon icon (32x32)
│   ├── icon-64.png     # High-res icon
│   └── icon-80.png     # High-res icon
└── README.md           # This file
```

## Development

### Local Testing

For local development, you can use a simple HTTP server:

```bash
# Using Python 3
python -m http.server 3000

# Using Node.js (npx)
npx http-server -p 3000
```

Then update the manifest URLs to use `https://localhost:3000/...`

Note: You'll need HTTPS for production. For local testing, you may need to configure Excel to trust localhost.

### Technologies Used

- **Office.js**: Microsoft's JavaScript API for Office Add-ins
- **HTML5/CSS3**: Modern web standards
- **Vanilla JavaScript**: No framework dependencies

## Credits

- Original DataGen VBA: Dr. Russell T. Hurlburt
- Office Add-in Implementation: dep2025

## License

Educational use. Based on DataGen by Dr. Russell Hurlburt, UNLV.

## Troubleshooting

### "Add-in not loading"
- Ensure your GitHub Pages site is published and accessible
- Check that all URLs in manifest.xml are correct
- Clear browser cache and try again

### "Statistics not updating"
- Click the "Refresh Statistics" button
- Ensure data is in columns B-F, rows 4-203
- Check that cells contain numeric values (not text)

### "CORS errors in console"
- This usually means the add-in files aren't being served with correct headers
- GitHub Pages should handle this automatically

## Future Enhancements

- [ ] Histogram visualization
- [ ] Export statistics to worksheet
- [ ] Custom data ranges
- [ ] Additional distributions (exponential, Poisson, etc.)
- [ ] ANOVA calculations
- [ ] Chi-square tests
