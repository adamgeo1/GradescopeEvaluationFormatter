from MidtermReader import get_midterm_data
from GradescopeReader import get_survey_data
from scipy.stats import chi2_contingency
import pandas as pd
from pathlib import Path

OUTPUT_PATH = Path(__file__).parent.absolute() / "output" / "ERsurvey_midterm_comparison.xlsx"

def main():
    midterm = get_midterm_data()
    survey = get_survey_data()

    midterm_qs = {
        'Q1' : 8,
        'Q2' : 20,
        'Q3' : 12,
        'Q4' : 40,
        'Q5_6' : 20
    }

    results = []

    for i in range(2): # tests with perfect scores and perfect - 1 scores

        if i != 0:
            results.append({
                "Survey Question": "",
                "Midterm Question": "",
                "Survey Correct & MT Correct": "",
                "Survey Correct & MT Incorrect": "",
                "Survey Incorrect & MT Correct": "",
                "Survey Incorrect & MT Incorrect": "",
                "Chi-squared": "",
                "p-value": ""
            })

            results.append({
                "Survey Question": f"Tests allowing for {i} less than perfect score:",
                "Midterm Question": "",
                "Survey Correct & MT Correct": "",
                "Survey Correct & MT Incorrect": "",
                "Survey Incorrect & MT Correct": "",
                "Survey Incorrect & MT Incorrect": "",
                "Chi-squared": "",
                "p-value": ""
            })

            results.append({
                "Survey Question": "",
                "Midterm Question": "",
                "Survey Correct & MT Correct": "",
                "Survey Correct & MT Incorrect": "",
                "Survey Incorrect & MT Correct": "",
                "Survey Incorrect & MT Incorrect": "",
                "Chi-squared": "",
                "p-value": ""
            })

        for midterm_q, score in midterm_qs.items():

            rubric_index = 0

            for survey_q in survey[next(iter(survey))].keys():  # Iterate through all survey questions
                a = b = c = d = 0

                for sid in survey:
                    if sid in midterm and midterm_q in midterm[sid] and survey_q in survey[sid]:
                        if rubric_index >= len(survey[sid][survey_q]) or rubric_index >= len(midterm[sid][midterm_q][1]):
                            continue

                        survey_correct = survey[sid][survey_q][rubric_index] == 1
                        midterm_correct = midterm[sid][midterm_q][0] >= score - i

                        if survey_correct and midterm_correct:
                            a += 1
                        elif survey_correct and not midterm_correct:
                            b += 1
                        elif not survey_correct and midterm_correct:
                            c += 1
                        else:
                            d += 1

                table = [[a, b], [c, d]]

                if a < 5 or b < 5 or c < 5 or d < 5:
                    print(f"{survey_q}: Skipping — not enough variation for chi-squared.\n")
                    results.append({
                        "Survey Question": survey_q,
                        "Midterm Question": midterm_q,
                        "Survey Correct & MT Correct": a,
                        "Survey Correct & MT Incorrect": b,
                        "Survey Incorrect & MT Correct": c,
                        "Survey Incorrect & MT Incorrect": d,
                        "Chi-squared": "Not enough variation for chi-squared.",
                        "p-value": ""
                    })
                    continue

                chi2, p, dof, expected = chi2_contingency(table)

                results.append({
                    "Survey Question": survey_q,
                    "Midterm Question": midterm_q,
                    "Survey Correct & MT Correct": a,
                    "Survey Correct & MT Incorrect": b,
                    "Survey Incorrect & MT Correct": c,
                    "Survey Incorrect & MT Incorrect": d,
                    "Chi-squared": round(chi2, 4),
                    "p-value": round(p, 4)
                })

    output_df = pd.DataFrame(results)
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        output_df.to_excel(writer, index=False)
        worksheet = writer.sheets["Sheet1"]

        from openpyxl.styles import PatternFill
        from copy import copy

        highlight = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

        for row_idx, (survey_q, p_val) in enumerate(zip(output_df["Survey Question"], output_df["p-value"]), start=2):
            # Bold header rows
            if isinstance(survey_q, str) and survey_q.startswith("Tests allowing for"):
                for col in range(1, len(output_df.columns) + 1):
                    cell = worksheet.cell(row=row_idx, column=col)
                    bold_font = copy(cell.font)
                    bold_font.bold = True
                    cell.font = bold_font

            # Highlight p-value < 0.05
            if isinstance(p_val, (int, float)) and p_val < 0.05:
                cell = worksheet.cell(row=row_idx, column=output_df.columns.get_loc("p-value") + 1)
                cell.fill = highlight

        for col_idx, column in enumerate(output_df.columns, start=1):
            header_text = str(column)
            adjusted_width = len(header_text) + 2  # slight padding
            col_letter = worksheet.cell(row=1, column=col_idx).column_letter
            worksheet.column_dimensions[col_letter].width = adjusted_width

    print(f"\nAll results written to {OUTPUT_PATH}")

if __name__ == "__main__":
    main()
