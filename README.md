# Excel Sheet Automation

This script automates the process of generating feedback for student grades in an Excel sheet using OpenAI's text generation capabilities. The script reads data from the Excel sheet, generates feedback based on the provided prompt, and fills in the corresponding cells in the sheet with the generated feedback.

## Prerequisites

- Python 3.x
- openpyxl library (`pip install openpyxl`)
- OpenAI Python library (`pip install openai`)

## Setup

1. Clone or download this repository to your local machine.
    
2. Install the required dependencies by running the following command:
```bash
pip install openpyxl openai
```
    
- Open the `main.py` file and modify the following variables according to your requirements:
    
    - `openai.api_key`: Set your OpenAI API key. You can obtain this key from your OpenAI account.
    - `sheet_name`: Specify the name of the sheet in your Excel file that contains the student grades.
    - `exclude_rows`: Specify the rows to exclude from generating feedback. These rows will be left empty.

## Usage

To use the script, follow these steps:

1. Open a command prompt or terminal.
    
2. Navigate to the directory where the script is located.
    
3. Run the script with the following command:
```bash
python automation.py [file_path] [start_column] [end_column] [start_row] [end_row]
```

Replace the placeholders `[file_path]`, `[start_column]`, `[end_column]`, `[start_row]`, and `[end_row]` with the appropriate values:

- `file_path`: The path to the Excel file.
- `start_column`: The column letter (uppercase) where the feedback should start populating (e.g., 'B').
- `end_column`: The column letter (uppercase) where the feedback should stop populating (e.g., 'F').
- `start_row`: The row number where the feedback should start populating.
- `end_row`: The row number where the feedback should stop populating.

Example command:

```bash
python automation.py excel_file.xlsx B F 3 49
```
    
4. The script will generate feedback for each student's grade in the specified range and fill in the corresponding cells in the Excel sheet.
    
5. Once the process is completed, the updated Excel file will be saved automatically.

Note: The script utilizes the OpenAI API, and there may be usage costs associated with making API calls.

## License

This project is licensed under the [MIT License](https://chat.openai.com/LICENSE).
