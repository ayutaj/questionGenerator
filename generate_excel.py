import pandas as pd
import argparse

def generate_excel(file_name: str):
    # Define required columns
    required_cols = [
        "Question No", "question",
        "option-a", "option-b", "option-c", "option-d",
        "answer"
    ]
    
    # Create empty DataFrame with only headers
    df = pd.DataFrame(columns=required_cols)
    
    # Save to Excel
    df.to_excel(file_name, index=False)
    print(f"Template created successfully: {file_name}")


def main():
    parser = argparse.ArgumentParser(description="Generate question template Excel file.")
    parser.add_argument("--name", required=True, help="Output Excel file name (e.g., questions.xlsx)")
    
    args = parser.parse_args()
    
    # Generate the excel file
    generate_excel(args.name)


if __name__ == "__main__":
    main()