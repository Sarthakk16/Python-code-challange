import pandas as pd
import re
from collections import Counter

def extract_groups_from_comments(comment):
    """Extract groups mentioned in the comment using regex."""
    match = re.search(r'Groups : \[code\]<I>(.*?)</I>\[/code\]', comment, re.IGNORECASE)
    if match:
        groups = match.group(1).split(',')  # Split multiple groups if separated by commas
        return [group.strip() for group in groups]
    return []

def process_groups_from_excel(file_path, sheet_name='Sheet1', column_name='Additional comments'):
    """Read the Excel file and process group occurrences."""
    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype=str)
    
    group_counter = Counter()
    for comment in df[column_name].dropna():
        groups = extract_groups_from_comments(comment)
        group_counter.update(groups)
    
    result_df = pd.DataFrame(group_counter.items(), columns=['Group name', 'Number of occurrences'])
    result_df.sort_values(by='Number of occurrences', ascending=False, inplace=True)
    return result_df

def save_results_to_csv(result_df, output_file='output.csv'):
    """Save the result to a CSV file."""
    result_df.to_csv(output_file, index=False)
    print(f'Results saved to {output_file}')

# Example usage
if __name__ == "__main__":
    input_file = 'Input-Data.xlsx'  # Change this to the actual file path
    output_csv = 'Output.csv'
    result_df = process_groups_from_excel(input_file)
    save_results_to_csv(result_df, output_csv)
