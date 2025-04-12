import pandas as pd
import os

def clean_excel_file(file_path, output_dir=None):
    # Load the Excel or CSV file
    if file_path.endswith('.csv'):
        df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)

    print(f"Original shape: {df.shape}")

    # Drop completely empty rows
    df.dropna(how='all', inplace=True)

    # Drop duplicate rows
    df.drop_duplicates(inplace=True)

    # Strip whitespace from column names and fix formatting
    df.columns = [col.strip().title().replace('_', ' ') for col in df.columns]

    # Strip whitespace from string-type values
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    print(f"Cleaned shape: {df.shape}")

    # Prepare output path
    filename = os.path.basename(file_path)
    name, ext = os.path.splitext(filename)
    cleaned_name = f"{name}_cleaned.csv"
    output_path = os.path.join(output_dir or os.getcwd(), cleaned_name)

    # Save cleaned version as CSV
    df.to_csv(output_path, index=False)
    print(f"Saved cleaned file: {output_path}")

# Example usage
if __name__ == "__main__":
    path = input("Enter path to your Excel or CSV file: ").strip()
    clean_excel_file(path)
