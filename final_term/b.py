import csv
import glob
from collections import defaultdict
import os

def parse_csv_for_math_stats(filepaths):
    """
    Parses multiple CSV files to gather math statistics for each school.

    Args:
        filepaths (list): A list of paths to the CSV files.

    Returns:
        dict: A dictionary where keys are school names and values are
              dictionaries containing 'assigned_tasks', 'completed_tasks',
              and 'total_correct_rate_sum', 'correct_rate_count'.
    """
    school_stats = defaultdict(lambda: {
        "assigned_tasks": 0,
        "completed_tasks": 0,
        "total_correct_rate_sum": 0.0,
        "correct_rate_count": 0  # To count valid rates for averaging
    })

    encodings_to_try = ['utf-8', 'big5', 'cp950', 'big5hkscs', 'utf-8-sig'] # Added 'big5hkscs' and 'utf-8-sig'

    for filepath in filepaths:
        filename = os.path.basename(filepath)
        print(f"Processing file: {filename}...")
        file_processed = False
        for encoding in encodings_to_try:
            try:
                with open(filepath, mode='r', newline='', encoding=encoding, errors='strict') as csvfile:
                    # Skip the first two header lines based on observed CSV structure
                    # We need to be careful if some files have different header structures
                    try:
                        next(csvfile) # Skip first header line
                        next(csvfile) # Skip second header line
                    except StopIteration:
                        print(f"Warning: File {filename} (with {encoding}) has less than 2 header lines. Skipping file.")
                        # break # from inner loop, try next file or handle as error
                        # continue # to next encoding for this file might be problematic if headers are truly missing
                        # For now, let's assume if headers are short, the file is problematic for current logic
                        raise Exception("Too few header lines")


                    reader = csv.reader(csvfile)
                    temp_data_for_file = [] # Store rows before confirming full read
                    line_num = 2 # Start after 2 skipped header lines

                    for row in reader:
                        line_num += 1
                        try:
                            if len(row) > 16:  # Ensure column 4 (school) and 17 (rate) exist
                                school_name = row[3].strip()
                                correct_rate_str = row[16].strip()

                                if not school_name: # Skip if school name is empty
                                    # print(f"Warning: Empty school name in {filename} at original row ~{line_num}, line: {'|'.join(row)}")
                                    continue

                                temp_data_for_file.append({'school': school_name, 'rate_str': correct_rate_str, 'line': line_num})
                            else:
                                print(f"Warning: Row with insufficient columns in {filename} (encoding {encoding}) at data row ~{line_num-2}. Row: {row}")
                        except IndexError:
                            print(f"Warning: IndexError processing row in {filename} (encoding {encoding}) at data row ~{line_num-2}. Row: {row}")
                        except Exception as e_row: # Catch other errors during row processing
                            print(f"Error processing a row in {filename} (encoding {encoding}) at data row ~{line_num-2}: {e_row}. Row: {row}")
                    
                    # If we successfully read all rows with this encoding
                    for data_item in temp_data_for_file:
                        school_name = data_item['school']
                        correct_rate_str = data_item['rate_str']
                        
                        school_stats[school_name]["assigned_tasks"] += 1
                        try:
                            correct_rate = float(correct_rate_str)
                            # Assuming a task is "completed" if the rate is a valid number
                            # You might have a different definition of "completed"
                            school_stats[school_name]["completed_tasks"] += 1
                            school_stats[school_name]["total_correct_rate_sum"] += correct_rate
                            school_stats[school_name]["correct_rate_count"] += 1
                        except ValueError:
                            # If correct_rate_str is not a valid float, it's an assigned task but not "completed" with a valid rate
                            # print(f"Notice: Non-numeric or empty correct rate for {school_name} in {filename} (original row ~{data_item['line']}): '{correct_rate_str}'")
                            pass # Still counts as assigned

                print(f"Successfully processed {filename} with encoding {encoding}")
                file_processed = True
                break  # Break from encodings loop since file was processed
            except UnicodeDecodeError:
                # print(f"Failed to decode {filename} with {encoding}...")
                pass # Try next encoding
            except FileNotFoundError:
                print(f"Error: File {filename} not found.")
                break # Break from encodings loop, no point trying other encodings
            except Exception as e:
                print(f"An unexpected error occurred with {filename} (encoding {encoding}): {e}")
                # Depending on the error, you might want to break or continue
                # If it's an error like "Too few header lines", we might want to stop for this file.
                # For now, let's try the next encoding if it's not a FileNotFoundError
                if isinstance(e, FileNotFoundError): # Should be caught above, but for safety
                    break
        
        if not file_processed:
            print(f"Error: Could not decode or process {filename} with any attempted encodings. It will be skipped.")

    return school_stats

def print_stats_table(school_stats):
    """
    Prints the aggregated statistics in a formatted table.
    """
    print("\n--- 數學任務統計報告 ---")
    header = "| {:<20} | {:<15} | {:<15} | {:<15} | {:<15} |".format(
        "學校名稱", "總派出任務數", "總已完成任務數", "整體完成率", "平均正答率"
    )
    print(header)
    print("-" * len(header))

    # Sort by school name for consistent output, if desired
    # sorted_school_names = sorted(school_stats.keys())
    # For now, use the order they were encountered or defaultdict's internal order
    
    for school_name, stats in school_stats.items():
        assigned = stats['assigned_tasks']
        completed = stats['completed_tasks']
        
        if assigned > 0:
            completion_rate = (completed / assigned) * 100
        else:
            completion_rate = 0.0

        if stats['correct_rate_count'] > 0: # Use correct_rate_count for average
            avg_correct_rate = (stats['total_correct_rate_sum'] / stats['correct_rate_count']) * 100
            avg_correct_rate_str = f"{avg_correct_rate:.2f}%"
        else:
            avg_correct_rate_str = "N/A"

        print("| {:<20} | {:<15} | {:<15} | {:<15.2f}% | {:<15} |".format(
            school_name,
            assigned,
            completed,
            completion_rate,
            avg_correct_rate_str
        ))
    print("-" * len(header))
    print("\n註：")
    print("1. 總派出任務數：該學校在所有提供的年級CSV檔案中出現的總行數（每行視為一筆任務記錄）。")
    print("2. 總已完成任務數：派出任務中，其『整體正答率』欄位（第17欄）為有效數值的任務數。")
    print("3. 整體完成率 = (總已完成任務數 / 總派出任務數) * 100%。")
    print("4. 平均正答率：所有已完成任務的『整體正答率』之平均值。")
    print("5. 學校名稱的呈現依賴於CSV檔案中的編碼與內容。若有亂碼，表示編碼問題或檔案本身問題。")


def main():
    """
    Main function to orchestrate the CSV parsing and printing of stats.
    """
    # If the script is inside the 'final' directory,
    # CSV files are in the current directory.
    csv_files = glob.glob("*.csv") # Changed from "final/*.csv"

    # You might want to add a check to ensure files are found
    if not csv_files:
        print("No CSV files found in the current directory.")
        # Attempt to find files if the script is run from one level above 'final'
        # This makes the script more robust to where it's run from.
        print("Attempting to find CSV files in a 'final' subdirectory...")
        csv_files = glob.glob("final/*.csv")
        if not csv_files:
            print("No CSV files found in 'final' subdirectory either.")
            # For debugging, print current working directory and its contents
            current_path = os.getcwd()
            print(f"Current working directory: {current_path}")
            try:
                print(f"Files in '{current_path}': {os.listdir(current_path)}")
            except FileNotFoundError:
                print(f"Directory '{current_path}' not found or inaccessible.")
            
            base_path_for_final_check = os.path.join(current_path, "final")
            if os.path.isdir(base_path_for_final_check):
                 print(f"Files in '{base_path_for_final_check}': {os.listdir(base_path_for_final_check)}")
            else:
                print(f"Subdirectory 'final' ({base_path_for_final_check}) does not exist from current location.")
            return

    print(f"Found files: {csv_files}")
    
    aggregated_stats = parse_csv_for_math_stats(csv_files)
    
    if aggregated_stats:
        print_stats_table(aggregated_stats)
    else:
        print("No statistics were generated. Please check file contents and processing logs.")

if __name__ == "__main__":
    main()
