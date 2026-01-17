# DTS-SAW Excel to XML Converter

A lightweight Python tool designed to convert manufacturing data from Excel spreadsheets (`.xlsx`) into XML files compatible with industrial saw cutting machines.

## Features

- **Batch Processing**: Converts every Excel file in the directory at once.
- **Template Support**: Tailored for specific panel cutting layouts.
- **Intelligent Grain Handling**: Supports numeric (1/0) or text (yes/no) grain specifications.
- **Secondary Descriptions**: Automatically detects and includes optional secondary descriptions if provided.
- **Windows Friendly**: Includes a one-click `.bat` script for non-technical users.

---

## Excel Template Specification

The program expects the following column layout starting from the first sheet:

| Column | Name | Description |
| :--- | :--- | :--- |
| **A** | - | *Ignored* |
| **B** | - | *Ignored* |
| **C** | **Length** | Panel length (dimension along grain) |
| **D** | **Width** | Panel width (dimension across grain) |
| **E** | **Quantity** | Number of parts to cut |
| **F** | **Grain** | `1` or `yes` for grain tracking, `0` or `no` for none |
| **G** | **Description** | Main part description |
| **H** | **Sec. Desc** | *Optional* secondary description |

---

## Installation & Setup

### For Windows Users (Non-Technical)
1. **Install Python**: Download from [python.org](https://www.python.org/downloads/). 
   - **Important**: Check the box **"Add Python to PATH"** during installation.
2. **Download Files**: Ensure `dts-saw.py`, `requirements.txt`, and `run_conversion.bat` are in the same folder.
3. **Run**: Double-click `run_conversion.bat`. It will handle dependency installation and conversion automatically.

### For Developers / MacOS
1. **Clone the repo**:
   ```bash
   git clone https://github.com/waldekglaz/dts-saw-xml-converter.git
   cd dts-saw-xml-converter
   ```
2. **Setup Environment**:
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```
3. **Run**:
   ```bash
   python dts-saw.py
   ```

---

## Usage

1. Place your `.xlsx` files into the same folder as the script.
2. Run the script (via `.bat` file or terminal).
3. The program will generate a `.xml` file for each Excel file found, using the same filename.
4. Error logs will be displayed in the console if any rows contain invalid data.

## Output Format

The generated XML follows the `<CutList>` schema required by industrial optimization software, including predefined machine parameters, blade thicknesses, and material codes (default: 18mm MDF).

---

## License
MIT
