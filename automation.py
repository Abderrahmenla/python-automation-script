import sys
from openpyxl import load_workbook
import time
import openai

openai.api_key = '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'


def send_prompt_to_chat(prompt):
    try:
        response = openai.Completion.create(
            engine='text-davinci-003',
            prompt=prompt,
            max_tokens=100,
            n=1,
            stop=None,
            temperature=0.7,
            top_p=1.0,
            frequency_penalty=0.0,
            presence_penalty=0.0
        )
        generated_text = response.choices[0].text.strip()
        return generated_text
    except Exception as e:
        print(f"Error generating chat response: {e}")
        return ""

def parse_excel_sheet(file_path, sheet_name):
    try:
        workbook = load_workbook(file_path)
        sheet = workbook[sheet_name]
        start_row = 3
        end_row = 49
        start_column = ord('B') - 64
        end_column = ord('F') - 64
        exclude_rows = [8, 9, 18, 19, 25, 26, 31, 32, 36, 37, 42, 43]
        data = []

        total_rows = end_row - start_row + 1
        processed_rows = 0

        for row_num, row in enumerate(sheet.iter_rows(min_row=start_row, max_row=end_row, min_col=start_column, max_col=end_column, values_only=True), start=start_row):
            if row_num in exclude_rows:
                data.append([])
                continue

            row_data = []
            for i in reversed(range(1, len(row))):
                if i == 0:
                    break 
                cell_value = row[i]
                next_column = i + 1
                next_cell_value = sheet.cell(row=row_num, column=next_column).value
                prompt = f"A student got a grade with this description: {cell_value} Write a feedback on how he can improve and get the next grade which is: {next_cell_value}"
                generated_feedback = send_prompt_to_chat(prompt)
                time.sleep(10)
                row_data.append(generated_feedback)
            
            data.append(row_data)

            processed_rows += 1
            print(f"Processed {processed_rows} out of {total_rows} rows.")

        return data
    except Exception as e:
        print(f"Error parsing Excel sheet: {e}")
        return []

def fill_excel_with_data(file_path, sheet_name, data, start_column, end_column, start_row, end_row, exclude_rows):
    try:
        wb = load_workbook(filename=file_path)
        sheet = wb[sheet_name]

        for row_index, row in enumerate(data):
            if row_index + start_row in exclude_rows:
                continue

            for column_index, value in enumerate(reversed(row)):
                column = start_column + column_index
                cell = sheet.cell(row=start_row+row_index, column=column)
                cell.value = value

                if column == end_column:
                    break

                if start_row+row_index == end_row:
                    break

        wb.save(file_path)
        print(f"Data saved to {file_path}")
    except Exception as e:
        print(f"Error filling Excel with data: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 6:
        print("Please provide the file path, start column, end column, start row, and end row as command-line arguments.")
        sys.exit(1)

    file_path = sys.argv[1]
    sheet_name = "Criteria Classifications"  # Specify the desired sheet name here
    start_column = ord(sys.argv[2].upper()) - 64
    end_column = ord(sys.argv[3].upper()) - 64
    start_row = int(sys.argv[4])
    end_row = int(sys.argv[5])

    data = parse_excel_sheet(file_path, sheet_name)

    exclude_rows = [8, 9, 18, 19, 25, 26, 31, 32, 36, 37, 42, 43]

    fill_excel_with_data(file_path, sheet_name, data, start_column, end_column, start_row, end_row, exclude_rows)
