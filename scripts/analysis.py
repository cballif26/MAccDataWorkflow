import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import textwrap
import os

# Ensure outputs directory exists
os.makedirs('outputs', exist_ok=True)

def main():
    # 1. Read Data
    # Header is in row 0 (0-indexed).
    # Row 1 (index 0 in df) is Question Text.
    # Row 2 (index 1 in df) is Metadata.
    # Data starts at row 3 (index 2 in df).
    try:
        # Load all as object first to avoid type inference issues on header rows
        df = pd.read_excel('data/Grad Program Exit Survey Data 2024.xlsx', header=0)
    except FileNotFoundError:
        print("Error: File not found.")
        return

    # Extract Question Texts (Row index 0 in dataframe, which is Row 1 in Excel)
    question_texts = df.iloc[0]

    # Extract Data (Row index 2 onwards in dataframe, which is Row 3+ in Excel)
    # We copy to avoid SettingWithCopy warnings
    data = df.iloc[2:].copy()

    # List to store processed dataframes
    melted_dfs = []

    # --- 2. Process Q35 Columns (Core Courses - Ranks 1-8) ---
    # Question: "Place each MAcc CORE course into rank order..."
    # Logic: Rank 1 is Best, Rank 8 is Worst.
    # Target: Rating 5 (Best) to 1 (Worst).
    # Formula: Rating = 5 - ((Rank - 1) * (4/7))

    q35_cols = [c for c in df.columns if c.startswith('Q35_')]
    if q35_cols:
        # Convert to numeric, coercing errors
        for col in q35_cols:
            data[col] = pd.to_numeric(data[col], errors='coerce')

        # Filter columns that are actually numeric (in case of empty cols)
        # and normalize
        # We process column by column to handle NaNs correctly
        for col in q35_cols:
            # Normalize: 1 -> 5, 8 -> 1
            # Check if max is roughly 8 to be safe?
            # Assuming standard 8 items as per inspection.
            data[col] = 5 - ((data[col] - 1) * (4/7))

        # Melt
        q35_melt = data[q35_cols].melt(var_name='QuestionCode', value_name='Rating')

        # Extract Course Names
        def extract_course_name_q35(code):
            text = str(question_texts[code])
            if ' - ' in text:
                return text.split(' - ')[-1].strip()
            return text

        q35_melt['Course or Program Name'] = q35_melt['QuestionCode'].apply(extract_course_name_q35)

        # Drop NaNs
        q35_melt = q35_melt.dropna(subset=['Rating'])

        melted_dfs.append(q35_melt[['Course or Program Name', 'Rating']])

    # --- 3. Process Q76-Q83 Columns (Elective/Other - Ratings 1-5) ---
    # Question: "Rate ... on a scale from 1-5..."
    # Logic: 5 is Best. Keep as is.

    # Identify relevant columns based on text "Rate" and scale "1-5"
    rate_cols = []
    for col in df.columns:
        # Check if column name starts with Q76-Q83 (range based on inspection)
        # Actually inspection showed Q76..Q83.
        # Let's be broader: Check if question text contains "Rate" and "scale".
        q_text = str(question_texts[col])
        if "Rate" in q_text and "scale" in q_text:
            rate_cols.append(col)

    if rate_cols:
        for col in rate_cols:
            data[col] = pd.to_numeric(data[col], errors='coerce')

        rate_melt = data[rate_cols].melt(var_name='QuestionCode', value_name='Rating')

        def extract_course_name_rate(code):
            text = str(question_texts[code])
            # Text format: "Rate [Name] on a scale... - [Name]"
            if ' - ' in text:
                return text.split(' - ')[-1].strip()
            return text

        rate_melt['Course or Program Name'] = rate_melt['QuestionCode'].apply(extract_course_name_rate)

        # Drop NaNs
        rate_melt = rate_melt.dropna(subset=['Rating'])

        melted_dfs.append(rate_melt[['Course or Program Name', 'Rating']])

    # --- 4. Process Q58 Columns (Program Aspects - Agreement) ---
    # Question: "Please indicate the extent you agree..."
    # Logic: Map strings to 1-5.

    q58_cols = [c for c in df.columns if c.startswith('Q58')]
    if q58_cols:
        mapping = {
            'Strongly Agree': 5,
            'Agree': 4,
            'Neither agree nor disagree': 3,
            'Neither disagree or agree': 3,
            'Disagree': 2,
            'Strongly Disagree': 1
        }

        # Need to work on a copy for mapping
        q58_data = data[q58_cols].copy()
        for col in q58_cols:
            q58_data[col] = q58_data[col].map(mapping)
            q58_data[col] = pd.to_numeric(q58_data[col], errors='coerce')

        q58_melt = q58_data.melt(var_name='QuestionCode', value_name='Rating')

        def extract_program_name_q58(code):
            text = str(question_texts[code])
            if ' - ' in text:
                return text.split(' - ')[-1].strip()
            return text

        q58_melt['Course or Program Name'] = q58_melt['QuestionCode'].apply(extract_program_name_q58)

        # Drop NaNs
        q58_melt = q58_melt.dropna(subset=['Rating'])

        melted_dfs.append(q58_melt[['Course or Program Name', 'Rating']])

    # --- 5. Combine and Analyze ---
    if not melted_dfs:
        print("No relevant data found.")
        return

    final_df = pd.concat(melted_dfs, ignore_index=True)

    # Calculate Mean Rating per Course/Program
    # We group by Name.
    ranking = final_df.groupby('Course or Program Name')['Rating'].mean().reset_index()

    # Sort Highest to Lowest (Rating descending), then by Name (ascending) for determinism
    ranking = ranking.sort_values(by=['Rating', 'Course or Program Name'], ascending=[False, True])

    # Save CSV
    ranking.to_csv('outputs/ranked_data.csv', index=False)
    print("Saved outputs/ranked_data.csv")

    # --- 6. Visualization ---
    # Wrap long names
    ranking['Wrapped Name'] = ranking['Course or Program Name'].apply(lambda x: '\n'.join(textwrap.wrap(str(x), 50)))

    # Set plot style
    sns.set_theme(style="whitegrid")

    # Figure size: height depends on number of items
    plt.figure(figsize=(10, max(6, len(ranking) * 0.5)))

    barplot = sns.barplot(data=ranking, x='Rating', y='Wrapped Name', color='skyblue')

    # Add sample size to title
    # We use the length of the data dataframe which corresponds to the number of respondents
    total_sample_size = len(data)
    plt.title(f'Course and Program Ratings (N={total_sample_size})')
    plt.xlabel('Mean Rating (1-5 Scale)')
    plt.ylabel('Course / Program')

    # Tight layout
    plt.tight_layout()

    plt.savefig('outputs/program_rankings.png')
    print("Saved outputs/program_rankings.png")

if __name__ == "__main__":
    main()
