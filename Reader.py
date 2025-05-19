import pandas as pd
from pathlib import Path

INPUT_PATH = Path(__file__).parent.absolute() / 'ERsurvey.csv'
OUTPUT_PATH = Path(__file__).parent.absolute() / 'ERsurveyFormat.csv'

def main():
    df = pd.read_csv(INPUT_PATH)
    ids_and_answers = {
        0 if pd.isna(k) else int(k): v
        for k, v in df.set_index('AnonID')['Q1: columns'].items()
    }

    questions = df.columns.tolist()
    questions = [e for e in questions if e[0] == 'Q']
    formatDF = pd.DataFrame({'AnonID': ids_and_answers.keys()})
    for q in questions:
        formatDF[q] = pd.NA
    formatDF.to_csv(OUTPUT_PATH, index=False)
    for id in ids_and_answers:
        print(f"{id}: {ids_and_answers[id]}")




if __name__ == '__main__':
    main()