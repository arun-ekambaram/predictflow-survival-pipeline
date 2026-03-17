df = df.reset_index(drop=True)

mask = df['run_date'].isna()
cols = df.columns.tolist()
shift_start = cols.index('')

for idx in df.index[mask]:
    row = df.loc[idx].tolist()

    # 1. merge split column
    row[shift_start] = str(row[shift_start]) + ', ' + str(row[shift_start + 1])

    # 2. REMOVE the extra column (this is key 🔥)
    del row[shift_start + 1]

    # 3. pad to maintain length
    row.append(None)

    df.loc[idx] = row
