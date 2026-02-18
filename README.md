# MAccDataWorkflow

## Automated Ranking Workflow

This repository includes a deterministic data analysis workflow that automatically ranks courses and program aspects based on student exit survey data.

### Methodology

The analysis is performed by `scripts/analysis.py`. It processes the data in `data/Grad Program Exit Survey Data 2024.xlsx` using the following steps:

1.  **Data Ingestion**: Reads the Excel file, using the first row as headers and properly handling question text and metadata rows.
2.  **Column Identification**:
    *   **Core Courses (Rankings)**: Identifies columns where students ranked courses (Q35 series). The ranks (1-8, where 1 is best) are inverted and normalized to a 1-5 scale to be comparable with other ratings. The formula used is: `Rating = 5 - ((Rank - 1) * (4/7))`. This ensures that a "Best" rank of 1 becomes a Rating of 5.
    *   **Elective Courses (Ratings)**: Identifies columns where students rated courses on a 1-5 scale (Q76-Q83 series). These are kept as-is.
    *   **Program Aspects**: Identifies columns where students indicated agreement (Q58 series). Responses are mapped to a numeric scale (Strongly Agree = 5, ..., Strongly Disagree = 1).
3.  **Data Transformation**: The data is "melted" into a long format containing "Course or Program Name" and "Rating".
4.  **Aggregation**: Calculates the mean rating for each course/program name.
5.  **Sorting**: Sorts the results by Mean Rating (Descending) and then by Name (Ascending) to ensure a definitive and deterministic rank order.
6.  **Visualization**: Generates a horizontal bar chart (`outputs/program_rankings.png`) displaying the rankings. Long names are wrapped for readability.
7.  **Output**: Saves the ranked data to `outputs/ranked_data.csv`.

### Automation

A GitHub Actions workflow (`.github/workflows/main.yml`) is configured to run on every push to the `main` branch. This workflow:
1.  Sets up a Python environment.
2.  Installs dependencies from `requirements.txt`.
3.  Executes the analysis script.
4.  Commits and pushes any updated results in the `outputs/` folder back to the repository.

This ensures that the rankings and charts are always up-to-date with the latest code or data changes.
