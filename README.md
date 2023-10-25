# Hikmicro-Temperature-Matrix-Convert-CSV-to-xlsx

```
import pandas as pd
import os

def determine_delimiter(line):
    """
    Determine the delimiter used in a line.
    """
    if ';' in line and ',' in line:
        return ',' if line.count(',') > line.count(';') else ';'
    elif ';' in line:
        return ';'
    elif ',' in line:
        return ','
    else:
        return None

def convert_csv_to_xlsx_manual_floats_updated(csv_path):
    """
    Manually parse the .csv file, convert matrix data (from row 17 onwards) to floats, 
    and save it as .xlsx.
    """
    lines_list = []
    current_row = 1  # To keep track of the row number
    
    with open(csv_path, 'r', encoding='utf-8') as file:
        for line in file:
            delim = determine_delimiter(line.strip())
            if delim:
                split_line = line.strip().split(delim)
                
                # Blank out the information for rows from 1 to 16
                if 1 <= current_row <= 16:
                    split_line = [''] * len(split_line)
                
                # Convert to float for rows from 17 onwards
                elif current_row >= 17:
                    for idx, item in enumerate(split_line):
                        try:
                            split_line[idx] = float(item)
                        except ValueError:
                            # If conversion to float is not possible, keep it as is
                            pass
                
                lines_list.append(split_line)
            current_row += 1

    # Convert the list of lists into a DataFrame
    df = pd.DataFrame(lines_list)

    # Save the DataFrame to an .xlsx file
    xlsx_path = csv_path.replace('.csv', '.xlsx')
    with pd.ExcelWriter(xlsx_path, engine='openpyxl') as excel_writer:
        df.to_excel(excel_writer, index=False, header=False)

    return xlsx_path


if __name__ == "__main__":
    # List all files in the current directory with .csv extension
    csv_files = [f for f in os.listdir() if f.endswith('.csv')]

    # Convert each .csv file to .xlsx
    for csv_file in csv_files:
        convert_csv_to_xlsx_manual_floats_updated(csv_file)
        print(f"Converted {csv_file} to XLSX.")
```
