from pathlib import Path
import pandas as pd

INPUT_PATH = Path(__file__).parent.absolute() / "Copy of midtermAnon - full.csv"

def safe_int(x):
    return 0 if pd.isna(x) or (isinstance(x, float) and math.isnan(x)) else int(x)

def main():
    df = pd.read_csv(INPUT_PATH)

    # The column that stores the student ID
    id_column = df.columns[0]

    # Define score and rubric column mappings for each question
    question_info = {
        "Q1": {"score": 1, "rubric": [2, 3, 4, 5, 6, 7, 8]},
        "Q2a": {"score": 10, "rubric": [11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]},
        "Q2b": {"score": 24, "rubric": [25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35]},
        "Q2c": {"score": 37, "rubric": [38, 39, 40, 41, 42, 43, 44]},
        "Q2d": {"score": 46, "rubric": [47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57]},
        "Q3": {"score": 59, "rubric": [60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81]},
        "Q4a": {"score": 83, "rubric": [84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100]},
        "Q4b": {"score": 102, "rubric": [103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115]},
        "Q4c": {"score": 117, "rubric": [118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131]},
        "Q4d": {"score": 133, "rubric": [134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144, 145, 146, 147, 148, 149, 150]},
        "Q5": {"score": 151, "rubric": [152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162]},
        "Q6": {"score": 164, "rubric": [165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176]},
    }

    # Collect all necessary columns
    needed_columns = [id_column]
    for q in question_info.values():
        needed_columns.append(df.columns[q["score"]])
        needed_columns.extend(df.columns[i] for i in q["rubric"])

    # Build the nested dictionary
    ids_and_scores = {}
    for row in df.loc[1:, needed_columns].itertuples(index=False, name=None):
        student_id = row[0]
        values = row[1:]

        question_dict = {}
        offset = 0
        for qname, qdata in question_info.items():
            score = int(values[offset])
            rubric = [safe_int(v) for v in values[offset + 1: offset + 1 + len(qdata["rubric"])]]
            question_dict[qname] = (score, rubric)
            offset += 1 + len(qdata["rubric"])

        ids_and_scores[student_id] = question_dict

    # Print the result
    for sid, questions in ids_and_scores.items():
        print(f"{sid}:")
        for q, (score, rubric) in questions.items():
            print(f"  {q}: ({score}, {rubric})")

if __name__ == "__main__":
    main()
