import pandas as pd
from pathlib import Path

# ---- CONFIG ----
INPUT_FILE  = Path(__file__).parent / "ERsurvey.csv"
OUTPUT_DIR  = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = OUTPUT_DIR / "ERsurveyFormat.xlsx"

# A simple palette; will cycle if you have >10 rubric items
COLORS = [
    "#4F81BD", "#C0504D", "#9BBB59", "#8064A2",
    "#4BACC6", "#F79646", "#92BBD9", "#FF9DAE",
    "#FFFF99", "#A6A6A6"
]

# ---- 1) Read the raw CSV without header, grab rows 0–1 as our headers ----
raw = pd.read_csv(INPUT_FILE, header=None)
header0 = raw.iloc[0].tolist()
header1 = raw.iloc[1].tolist()
data    = raw.iloc[2:].reset_index(drop=True)

# Build a MultiIndex DataFrame
df = data.copy()
df.columns = pd.MultiIndex.from_arrays([header0, header1])

lvl0 = df.columns.get_level_values(0).astype(str)
lvl1 = df.columns.get_level_values(1).astype(str)

# Find where each question begins (level0 like "Q2: …")
q_starts = [i for i,h in enumerate(lvl0) if h.strip().startswith("Q") and ":" in h]

# ---- 2) Write one "Summary" sheet with tables + charts ----
with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
    workbook  = writer.book
    worksheet = workbook.add_worksheet("Summary")
    writer.sheets["Summary"] = worksheet

    row = 0
    for qi, startcol in enumerate(q_starts):
        endcol = q_starts[qi+1] if qi+1 < len(q_starts) else len(lvl0)

        # 2a) Grab **all** rubric labels in [startcol..endcol)
        labels = []
        cols   = []
        for j in range(startcol, endcol):
            lab = lvl1[j].strip()
            if lab:
                labels.append(lab)
                cols.append(j)

        # 2b) Count the 1's (fill NaN→0 first)
        counts = [
            int(df.iloc[:, j].fillna(0).astype(int).sum())
            for j in cols
        ]

        # 2c) Write the table
        short = lvl0[startcol].split(":")[0].strip()
        worksheet.write(row,    0, f"{short} Response Counts")
        worksheet.write(row+1,  0, "Label")
        worksheet.write(row+1,  1, "Count")
        for k, (lab, cnt) in enumerate(zip(labels, counts)):
            worksheet.write(row+2+k, 0, lab)
            worksheet.write(row+2+k, 1, cnt)

        # 3) Build a column chart: one series per rubric item
        chart = workbook.add_chart({"type": "column"})
        for i, lab in enumerate(labels):
            chart.add_series({
                "name":     lab,
                "values":   ["Summary", row+2+i, 1, row+2+i, 1],
                "fill":     {"color": COLORS[i % len(COLORS)]},
            })

        # Style the chart
        chart.set_title   ({"name": short})
        chart.set_x_axis  ({"name": "Option", "label_position": "none"})
        chart.set_y_axis  ({"name": "Count", "major_gridlines": {"visible": True}})
        chart.set_legend  ({"position": "right"})

        # Dynamically size it so the legend fits
        # base height 150px + 20px per label
        max_label_len = max(len(lab) for lab in labels)
        chart_width = 480 + max_label_len * 7 # 7px per char
        chart_height = 150 + 20 * len(labels)
        chart.set_size({
            "width":  chart_width,
            "height": chart_height
        })

        # Insert next to the table
        worksheet.insert_chart(row+1, 3, chart, {"x_offset": 10, "y_offset": 10})

        # Advance for next question
        row += len(labels) + 5

print(f"✓ Written summary + charts to {OUTPUT_FILE}")
