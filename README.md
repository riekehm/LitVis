# LitVis

Welcome to **LitVis** – your new favorite tool for literature research!  
If you’ve ever found yourself drowning in endless papers and thought, "How should I organise all this information?" then LitVis might just be the lifesaver you need. This lightweight chart programme lets you import/export CSV files, edit cells with rich-text formatting (bold, italic, bullet lists, colors, etc.), apply conditional formatting, and perform advanced filtering. And the best part? It doesn’t automatically crunch numbers or format dates—so you won’t be annoyed by Excel’s overzealous auto-formatting!


## Disclaimer
LitVis is a hobby project created to assist with literature research management. It contains many (for me) useful features and I know that it works on my computer but if it doesn´t work on yours I am sorry, I can´t help :D 
But feel free to adjust or to improve the code! I am pretty sure I made stupid mistakes and left some unnessecary parts in the code but I am afraid to delete them.
!! LitVis does not automatically save changes - you have to do that by export project (to save the formatting as well) or export CSV (to e.g. open the chart in Excel later, but then formatting (colours etc. will not be exported)


## Features

- **Insert or delete Columns/Rows**

- **Rich‑Text Editing:**  
  Edit cell content with formatting options such as bold, italic, bullet lists, font colors, and sizes.

- **Conditional Formatting:**  
  Highlight specific keywords in your cells based on user‑defined rules.

- **Advanced Filtering:**  
  Filter table data with multiple conditions. You can also filter “in any column” by selecting the “Any Column” option.

- **CSV Import/Export:**  
  Easily import and export CSV files using a custom delimiter (e.g. “;” for Excel compatibility).

- **Project Saving/Loading:**  
  Save your current table (including column order, widths, and formatting rules) as a JSON project file and load it later.

- **Column and Row Visibility:**  
  Show or Hide Rows / Columns

## Installation & Usage

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/yourusername/LitVis.git
   cd LitVis

2. **Setup a Virtual Environment (optional but recommended):**
   
     ```bash
    python -m venv venv
     
    ### On Windows:
    venv\Scripts\activate
    ### On macOS/Linux:
    source venv/bin/activate

     
3. **Install Dependencies:**

     ```bash
    pip install PyQt5   
    
4. **Run the Application:**

    ```bash
    python LitVis.py

## Customization
**Formatting & Conditional Rules:**
Use the built-in dialogs to set rich-text formatting and define rules for highlighting specific words.

**Column Management:**
Hide or show columns using the dropdown menu in the "Other Functions" tab.

## Contributing
Feel free to fork this repository and submit pull requests. Please adhere to the coding style and include tests for any new features.

## License
This project is licensed under the MIT License.
