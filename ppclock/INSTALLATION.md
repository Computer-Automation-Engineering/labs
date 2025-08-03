# PPClock PowerPoint Add-in Installation Guide

## Quick Setup

### 1. Start the Development Server
```bash
# Navigate to the PPClock project directory
cd git/cae/labs/ppclock

# Start the development server
python3 -m http.server 8000
```

### 2. Install the Add-in in PowerPoint

#### For PowerPoint on Windows/Mac:
1. Open PowerPoint
2. Go to **Insert** > **Add-ins** > **Upload My Add-in**
3. Click **Browse** and select the `manifest.xml` file
4. Click **Upload**

#### For PowerPoint Online:
1. Open PowerPoint in your web browser
2. Go to **Insert** > **Add-ins** > **Upload My Add-in**
3. Upload the `manifest.xml` file

### 3. Using PPClock
1. Look for the **PPClock Timer** button in the **Home** tab
2. Click it to open the timer panel
3. Set your countdown time and click **Start Countdown**

## Development Testing

### Testing in PowerPoint
1. Make sure the web server is running (`python3 -m http.server 8000`)
2. The add-in should appear in the ribbon
3. Test all features:
   - Countdown timer
   - Pause/Resume
   - PowerPoint integration features
   - Slide navigation

### Debugging
- Open Developer Tools in PowerPoint (F12)
- Check the browser console for any errors
- Verify all files are being served correctly

## Files Structure
```
git/cae/labs/ppclock/
├── manifest.xml              # Add-in manifest
├── ppclock_addon.html        # PowerPoint add-in version
├── ppclock_web.html          # Standalone web version
├── ppclock.py               # Original Python version
├── timer.txt                # Project notes and log
├── assets/                  # Icons and resources
└── INSTALLATION.md          # This file
```

## Features

### Core Timer Features
- ✅ Countdown timer with minutes/seconds input
- ✅ Pause/Resume functionality
- ✅ Visual progress bar
- ✅ Professional UI design
- ✅ Keyboard shortcuts (Space = pause, Esc = stop)

### PowerPoint Integration
- ✅ Ribbon button integration
- ✅ Task pane interface
- ✅ Slide navigation controls
- ✅ Insert timer slide
- ✅ Start slideshow from add-in
- ✅ Get presentation information

### Advanced Features
- ✅ Real-time countdown display
- ✅ Color-coded time warnings
- ✅ Completion notifications
- ✅ Responsive design
- ✅ Cross-platform compatibility

## Troubleshooting

### Common Issues
1. **Add-in not appearing**: Verify the web server is running and manifest.xml is valid
2. **HTTPS errors**: For production, use HTTPS instead of HTTP
3. **Broken icons**: Add actual PNG icon files to the assets folder
4. **PowerPoint features not working**: Ensure you're testing in PowerPoint, not a regular browser

### Support
- Check the browser console for JavaScript errors
- Verify all file paths in manifest.xml are correct
- Ensure PowerPoint has permissions to load the add-in

## Next Steps for Production
1. Create professional icon files (16x16, 32x32, 80x80 PNG)
2. Deploy to HTTPS web server
3. Update manifest.xml with production URLs
4. Test across different PowerPoint versions
5. Submit to Microsoft AppSource (optional)

## Learning Outcomes
This project demonstrates:
- Office Add-in development with Office.js
- HTML/CSS/JavaScript integration with PowerPoint
- Professional UI design for productivity tools
- Cross-platform web application development
- Timer and threading concepts in JavaScript