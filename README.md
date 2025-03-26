# LitVis

LitVis is a lightweight chart programme designed for managing literature research.  
It enables you to import and export CSV files, edit cells with rich‑text formatting (bold, italic, bullet lists, colors, etc.), apply conditional formatting, advanced filtering, and save/load projects (including column order, widths, and custom rules).

##Disclaimer
LitVis is a hobby project created to assist with literature research management. While it contains many useful features, it is still under active development and may not cover all edge cases.


## Features

- **CSV Import/Export:**  
  Easily import and export CSV files using a custom delimiter (e.g. “;” for Excel compatibility).

- **Rich‑Text Editing:**  
  Edit cell content with formatting options such as bold, italic, bullet lists, font colors, and sizes.

- **Conditional Formatting:**  
  Highlight specific keywords in your cells based on user‑defined rules.

- **Advanced Filtering:**  
  Filter table data with multiple conditions. You can also filter “in any column” by selecting the “Any Column” option.

- **Project Saving/Loading:**  
  Save your current table (including column order, widths, and formatting rules) as a JSON project file and load it later.

- **Undo/Redo & Column Visibility Toggle:**  
  Easily undo changes and toggle column visibility through a dedicated dropdown menu.

## Installation & Usage

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/yourusername/LitVis.git
   cd LitVis

2. **Setup a Virtual Environment (optional but recommended):**
    python -m venv venv
  # On Windows:
  venv\Scripts\activate
  # On macOS/Linux:
  source venv/bin/activate
3. **Install Dependencies:**
    pip install PyQt5   
4. **Run the Application:**
    python LitVis.py

##Customization
**Formatting & Conditional Rules:**
Use the built-in dialogs to set rich-text formatting and define rules for highlighting specific words.

**Column Management:**
Hide or show columns using the dropdown menu in the "Other Functions" tab.

##Contributing
Feel free to fork this repository and submit pull requests. Please adhere to the coding style and include tests for any new features.

##License
This project is licensed under the MIT License.


