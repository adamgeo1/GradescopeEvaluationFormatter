import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
from collections import Counter

INPUT_PATH = Path(__file__).parent.absolute() / 'ERsurvey.csv'
OUTPUT_PATH = Path(__file__).parent.absolute() / 'ERsurveyFormat.csv'

def main():
    df = pd.read_csv(INPUT_PATH)
    # Drop the description row (row 0)
    df_clean = df.iloc[1:].copy()

    # Convert AnonID to int, replace NaNs with 0
    df_clean['AnonID'] = df_clean['AnonID'].apply(lambda x: 0 if pd.isna(x) else int(x))

    # Identify question start columns
    question_starts = [col for col in df.columns if isinstance(col, str) and col.startswith('Q')]

    # Create the result dictionary
    result = {}

    for _, row in df_clean.iterrows():
        anon_id = int(row['AnonID'])
        result[anon_id] = {}

        for q_start in question_starts:
            start_idx = df.columns.get_loc(q_start)
            next_q_idx = next(
                (df.columns.get_loc(col) for col in question_starts if df.columns.get_loc(col) > start_idx),
                len(df.columns)
            )
            question_values = row.iloc[start_idx:next_q_idx]

            # Convert to list of ints (skip NaN, leave as 0 or None if needed)
            clean_values = []
            for val in question_values:
                try:
                    clean_values.append(int(float(val)))
                except (ValueError, TypeError):
                    clean_values.append(0)

            result[anon_id][q_start] = clean_values

    formatDF = pd.DataFrame({'AnonID': result.keys()})

    for q in question_starts:
        formatDF[q] = formatDF['AnonID'].apply(
            lambda anon_id: (
                ''.join(str(i + 1) for i, val in enumerate(result[anon_id][q]) if val == 1)
                if anon_id in result and q in result[anon_id] else pd.NA
            )
        )

    formatDF.to_csv(OUTPUT_PATH, index=False)

    for q in question_starts:
        indices = ''.join(formatDF[q].dropna())
        counts = Counter(indices)
        x = sorted(counts.keys())
        y = [counts[k] for k in x]

        plt.bar(x, y)
        plt.xlabel('Rubric Index')
        plt.ylabel('Count')
        plt.title(f'{q} Response Distribution')
        plt.tight_layout()
        plt.show()

if __name__ == '__main__':
    main()