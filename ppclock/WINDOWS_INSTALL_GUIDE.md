# PPClock Installation Guide for Office 365 PowerPoint on Windows 11

## 📋 Before You Start

**What You'll Need:**
- Windows 11 computer with Office 365 PowerPoint
- PPClock files from USB drive
- Administrator permissions (may be required)
- 5-10 minutes for installation

---

## 🗂️ Step 1: Copy Files from USB Drive

### 1.1 Insert USB Drive
- Insert the USB drive containing PPClock into your computer
- Windows will automatically detect and open the drive

### 1.2 Copy PPClock Folder
**Screenshot Location:** *File Explorer showing USB drive contents*

1. Open **File Explorer** (Windows + E)
2. Navigate to your USB drive (usually appears as "Removable Disk")
3. You should see a folder named **"PPClock"**
4. **Right-click** on the PPClock folder
5. Select **"Copy"** from the context menu

### 1.3 Paste to Your Computer
**Screenshot Location:** *File Explorer showing Documents folder*

1. Navigate to your **Documents** folder (recommended location)
2. **Right-click** in an empty area
3. Select **"Paste"** from the context menu
4. Wait for all files to copy (should take 10-15 seconds)

### 1.4 Verify Files Copied
**Screenshot Location:** *PPClock folder contents showing all files*

Open the copied PPClock folder and verify you see these files:
- ✅ `manifest.xml`
- ✅ `ppclock-manifest.xml` (backup)
- ✅ `ppclock_addon.html`
- ✅ `ppclock_web.html`
- ✅ `start_server.bat`
- ✅ `USB_SETUP_GUIDE.md`
- ✅ `MANIFEST_HELP.txt`
- ✅ `assets` folder

---

## 🖥️ Step 2: Start the PPClock Server

### 2.1 Run Server Startup Script
**Screenshot Location:** *PPClock folder with start_server.bat highlighted*

1. In the PPClock folder, find **`start_server.bat`**
2. **Double-click** on `start_server.bat`

### 2.2 Windows Security Warning
**Screenshot Location:** *Windows Defender SmartScreen warning dialog*

If Windows shows a security warning:
1. Click **"More info"** (small link at bottom)
2. Click **"Run anyway"** button that appears
3. This is normal for custom scripts

### 2.3 Server Running Confirmation
**Screenshot Location:** *Command prompt window showing server running*

You should see a **black Command Prompt window** with text like:
```
Starting PPClock Server...
Server will be available at: http://localhost:8000
Serving HTTP on :: port 8000 (http://[::]:8000/) ...
```

**IMPORTANT:** Keep this window open! The server must run while using PPClock.

---

## 📊 Step 3: Install PPClock Add-in in PowerPoint

### 3.1 Open PowerPoint
**Screenshot Location:** *PowerPoint start screen*

1. Click the **Windows Start button**
2. Type **"PowerPoint"**
3. Click **"PowerPoint"** when it appears
4. Open a new blank presentation or existing presentation

### 3.2 Access Add-ins Menu
**Screenshot Location:** *PowerPoint ribbon showing Insert tab*

1. Click the **"Insert"** tab in the PowerPoint ribbon
2. Look for the **"Add-ins"** section on the right side of the ribbon
3. Click the **"Add-ins"** button (may show as an icon or dropdown)

### 3.3 Upload Custom Add-in
**Screenshot Location:** *Add-ins menu dropdown showing "Upload My Add-in" option*

1. From the Add-ins dropdown menu, select **"Upload My Add-in"**
2. A file browser dialog will open

### 3.4 Navigate to Manifest File
**Screenshot Location:** *File browser dialog showing PPClock folder*

1. Navigate to your **Documents** folder
2. Open the **PPClock** folder
3. Look for **`manifest.xml`**

### 3.5 Handle File Visibility Issues
**Screenshot Location:** *File browser with file type dropdown showing "All Files"*

**Can't see manifest.xml?** Try these solutions:

**Solution A - Change File Type Filter:**
1. In the file browser, look for a dropdown that says "Office Add-in Manifest (*.xml)"
2. Change it to **"All Files (*.*)"**
3. Now you should see `manifest.xml`

**Solution B - Use Backup File:**
1. Look for **`ppclock-manifest.xml`** instead
2. This is identical to manifest.xml but with a different name

**Solution C - Show File Extensions:**
1. In File Explorer, click **"View"** tab
2. Check the box for **"File name extensions"**
3. Return to PowerPoint file dialog

### 3.6 Select and Upload Manifest
**Screenshot Location:** *manifest.xml file selected in browser dialog*

1. Click on **`manifest.xml`** (or `ppclock-manifest.xml`)
2. Click **"Open"** button
3. PowerPoint will process the add-in (may take 10-15 seconds)

### 3.7 Installation Success
**Screenshot Location:** *PowerPoint ribbon showing PPClock Timer button in Home tab*

If successful, you should see:
1. A new **"PPClock Timer"** button in the **Home** tab ribbon
2. The button appears in the "Presentation Tools" group
3. No error messages displayed

---

## 🎯 Step 4: Test PPClock Functionality

### 4.1 Open PPClock
**Screenshot Location:** *PPClock Timer button highlighted in ribbon*

1. Click the **"PPClock Timer"** button in the Home ribbon
2. A task pane should open on the right side of PowerPoint

### 4.2 PPClock Interface
**Screenshot Location:** *PPClock task pane showing setup screen*

You should see the PPClock interface with:
- ✅ "PPClock" title
- ✅ "Connected to PowerPoint" status message
- ✅ Minutes input field
- ✅ Seconds input field
- ✅ Display Size dropdown
- ✅ "Start Countdown" button
- ✅ PowerPoint Integration controls

### 4.3 Test Basic Countdown
**Screenshot Location:** *PPClock setup with sample time entered*

1. Enter **"0"** in Minutes field
2. Enter **"10"** in Seconds field
3. Select your preferred **Display Size**
4. Click **"Start Countdown"**

### 4.4 Countdown Running
**Screenshot Location:** *PPClock showing active countdown with large timer display*

You should see:
- ✅ Large countdown display (10, 9, 8, 7...)
- ✅ Progress bar showing time remaining
- ✅ Pause and Stop buttons
- ✅ PowerPoint controls (Next Slide, Previous Slide, etc.)

### 4.5 Test Features
**Screenshot Location:** *PPClock showing different font sizes*

Test these features:
- ✅ **Pause/Resume**: Click Pause, then Resume
- ✅ **Font Sizes**: Try different display sizes
- ✅ **PowerPoint Controls**: Test Next/Previous slide buttons
- ✅ **Timer Completion**: Let countdown reach 00:00

---

## 🔧 Troubleshooting Common Issues

### Issue 1: "Can't find manifest.xml"
**Screenshot Location:** *File browser showing file type filter dropdown*

**Solution:**
1. Change file type filter to "All Files (*.*)"
2. Use `ppclock-manifest.xml` instead
3. Enable "Show file extensions" in Windows

### Issue 2: "Server not responding"
**Screenshot Location:** *Command prompt window showing server error*

**Solution:**
1. Check that `start_server.bat` is still running
2. Restart the server by double-clicking `start_server.bat` again
3. Verify no firewall blocking port 8000

### Issue 3: "Add-in won't install"
**Screenshot Location:** *PowerPoint security warning dialog*

**Solution:**
1. Check Office security settings
2. May need administrator permissions
3. Try closing and reopening PowerPoint

### Issue 4: "PPClock button missing"
**Screenshot Location:** *PowerPoint ribbon with no PPClock button*

**Solution:**
1. Check if add-in was uploaded successfully
2. Look in Insert > My Add-ins for installed add-ins
3. Try uploading manifest again

---

## ✅ Installation Complete!

### What You Can Do Now:
- ✅ Set countdown timers for presentations
- ✅ Choose from 4 different font sizes
- ✅ Use pause/resume functionality
- ✅ Control PowerPoint slides from the timer
- ✅ Insert timer slides into presentations

### Remember:
- Keep the Command Prompt window open while using PPClock
- The server must be running for the add-in to work
- You can minimize the Command Prompt but don't close it

---

## 📱 Alternative: Web Version

**Screenshot Location:** *Web browser showing ppclock_web.html*

If you can't install the PowerPoint add-in:
1. Open any web browser
2. Navigate to the PPClock folder
3. Double-click **`ppclock_web.html`**
4. Use the full-featured web timer with draggable popup

---

## 🆘 Need More Help?

Check these files in your PPClock folder:
- **`MANIFEST_HELP.txt`** - Specific manifest file solutions
- **`USB_SETUP_GUIDE.md`** - Quick setup reference
- **`timer.txt`** - Complete technical documentation

---

**PPClock Installation Guide**  
*Professional Countdown Timer for PowerPoint Presentations*  
✨ Compatible with Office 365 PowerPoint on Windows 11