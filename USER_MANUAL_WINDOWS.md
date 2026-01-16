# DTS-SAW Converter - Manual for Windows Users

This tool converts all Excel files (`.xlsx`) in a folder into XML files formatted for the saw machine.

## 1. First Time Setup (Do Once)

### Step A: Install Python
1. Go to [python.org/downloads](https://www.python.org/downloads/).
2. Click the yellow button **Download Python**.
3. Run the installer you downloaded.
4. **⚠️ CRITICAL:** On the first screen of the installer, you **MUST** check the box that says **"Add Python.exe to PATH"**.
   *(If you miss this, the tool will not work. You will need to uninstall and reinstall Python).*
5. Click **Install Now** and wait for it to finish.

### Step B: Prepare the Folder
1. Create a new folder on your Desktop (e.g., named `DTS-SAW`).
2. Copy the following 3 files into this folder:
   - `dts-saw.py`
   - `requirements.txt`
   - `run_conversion.bat`

---

## 2. How to Use (Daily Usage)

1. **Place your Excel files** into the `DTS-SAW` folder.
   - You can copy as many `.xlsx` files as you want.
2. **Double-click** the file **`run_conversion.bat`**.
   - A black window will pop up.
   - It will automatically install any needed software updates.
   - It will find all Excel files and convert them.
3. **Done!**
   - You will see new `.xml` files appear in the folder next to your Excel files.
   - You can read the screen to see if any errors occurred (like missing columns in the Excel file).


## 3. Manual Method (Advanced / If .bat fails)

If the `run_conversion.bat` file does not work for some reason, you can run the tool manually using the "Command Prompt".

1.  **Open the Folder**: Open the `DTS-SAW` folder where your files are.
2.  **Open Terminal**:
    *   Click on the **address bar** at the top of the folder window (where it says `> Desktop > DTS-SAW`).
    *   Type `cmd` inside that bar and press **Enter**.
    *   A black window ("Command Prompt") should open.
3.  **Install Requirements** (Only needed the first time):
    *   Type or copy-paste this command and hit Enter:
        ```bash
        pip install -r requirements.txt
        ```
4.  **Run the Tool**:
    *   Type this command and hit Enter:
        ```bash
        python dts-saw.py
        ```
5.  **View Results**:
    *   The screen will show you which files are being processed.
    *   When finished, you can close the black window.

## Troubleshooting

- **"Python is not recognized..."**
  - This means you didn't check the "Add to PATH" box when installing Python. Please reinstall Python and check that box.

- **The window closes instantly**
  - There might be a file missing. Make sure `dts-saw.py` and `requirements.txt` are in the folder.
