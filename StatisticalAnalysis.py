from MidtermReader import get_midterm_data
from GradescopeReader import get_survey_data
from scipy.stats import chi2_contingency

def main():
    midterm = get_midterm_data()
    survey = get_survey_data()

    midterm_q = "Q3"
    rubric_index = 0

    for survey_q in survey[next(iter(survey))].keys():  # Iterate through all survey questions
        a = b = c = d = 0

        for sid in survey:
            if sid in midterm and midterm_q in midterm[sid] and survey_q in survey[sid]:
                if rubric_index >= len(survey[sid][survey_q]) or rubric_index >= len(midterm[sid][midterm_q][1]):
                    continue

                survey_correct = survey[sid][survey_q][rubric_index] == 1
                midterm_correct = midterm[sid][midterm_q][0] == 12

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
            continue

        print(f"\nSurvey Q: {survey_q} vs Midterm Q3")
        print("Contingency Table:")
        print(f"        MT Correct | MT Incorrect")
        print(f"Survey Correct  {a:>5}          {b:>5}")
        print(f"Survey Incorrect{c:>5}          {d:>5}")

        chi2, p, dof, expected = chi2_contingency(table)
        print(f"Chi-squared: {chi2:.4f}, p-value: {p:.4f}, dof: {dof}")

if __name__ == "__main__":
    main()
