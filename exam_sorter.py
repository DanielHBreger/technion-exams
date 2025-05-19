import pandas as pd
from pyuca import Collator

threshold = 3

def load_xlsx(file_path):
    """
    Load an Excel file and return a DataFrame.
    """
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error loading file: {e}")
        return None

def replace_missing_exams(df):
    for col in ['מועד א\'', 'מועד ב\'']:
        df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

def get_maslulim(df):
    unique_maslulim = []
    for column in df.columns[2:9]:
        unique_maslulim.extend(df[column].unique())
    unique_maslulim = list(set([str(x).strip() for x in unique_maslulim]))
    unique_maslulim = [x for x in unique_maslulim if str(x) != 'nan']
    
    c = Collator("allkeys.txt")
    # Sort the unique maslulim using the Collator
    unique_maslulim.sort(key=c.sort_key)
    return unique_maslulim

def main():
    # Load the Excel file
    file_path = 'פיזיקה בחינות חורף לעבכוד רק עם הקובץ הזהתשפו.xlsx'
    df = load_xlsx(file_path)
    df = replace_missing_exams(df)
    unique_maslulim = get_maslulim(df)
    warnings = []
    warnings_by_maslul = {key:0 for key in unique_maslulim}
    date_cols = ['מועד א\'', 'מועד ב\'']
    with open('unique_maslulim.txt', 'w', encoding='utf-8') as f:
        indices = [0,1,9,10,11,12]
        interesting_columns = [df.columns[indice] for indice in indices]
        problems = []
        problems_maslul = []
        problems_before = []
        problem_days = []
        for maslul in unique_maslulim:
            df_maslul = df[df.apply(lambda row: row.astype(str).str.contains(maslul, case=False).any(), axis=1)][interesting_columns]
            df_no_na_a = df_maslul.dropna(subset=['מועד א\''])
            df_no_na_b = df_maslul.dropna(subset=['מועד ב\''])
            by_moed_a = df_no_na_a.sort_values(by=['מועד א\''])
            by_moed_b = df_no_na_b.sort_values(by=['מועד ב\''])
            diffs_a = by_moed_a['מועד א\''].diff()
            diffs_b = by_moed_b['מועד ב\''].diff()
            f.write(f"{maslul}:\n")
            f.write(f"  מועד א':\n")
            for idx in range(len(by_moed_a['שם קורס'])):
                if idx == 0:
                    f.write(f"    {by_moed_a['שם קורס'].iloc[idx]}: {by_moed_a['מועד א\''].iloc[idx]}\n")
                else:
                    if diffs_a.iloc[idx].days <= threshold:
                        problems.append(by_moed_a['שם קורס'].iloc[idx]+" מועד א'")
                        problems_maslul.append(maslul)
                        problems_before.append(by_moed_a['שם קורס'].iloc[idx-1] + " מועד א'")
                        problem_days.append(diffs_a.iloc[idx].days)
                    f.write(f"    {by_moed_a['שם קורס'].iloc[idx]}: {by_moed_a['מועד א\''].iloc[idx]} ({diffs_a.iloc[idx].days} ימים)\n")
            f.write(f"  מועד ב':\n")
            for idx in range(len(by_moed_b['שם קורס'])):
                if idx == 0:
                    f.write(f"    {by_moed_b['שם קורס'].iloc[idx]}: {by_moed_b['מועד ב\''].iloc[idx]}\n")
                else:
                    if diffs_b.iloc[idx].days <= threshold:
                        problems.append(by_moed_b['שם קורס'].iloc[idx]+" מועד ב'")
                        problems_maslul.append(maslul)
                        problems_before.append(by_moed_b['שם קורס'].iloc[idx-1] + " מועד ב'")
                        problem_days.append(diffs_b.iloc[idx].days)
                    f.write(f"    {by_moed_b['שם קורס'].iloc[idx]}: {by_moed_b['מועד ב\''].iloc[idx]} ({diffs_b.iloc[idx].days} ימים)\n")
    df_problems = pd.DataFrame(list(zip(problems_maslul, problems_before, problems, problem_days)), columns=['מסלול', 'קורס לפני', 'קורס', 'מספר ימים'])
    df_problems = df_problems.iloc[:,::-1]
    df_problems =  df_problems.sort_values(by=['קורס'])
    with open('warnings.md', 'w', encoding='utf-8') as f:
        f.write(df_problems.to_markdown(index=False, tablefmt='pipe', colalign=['center']*len(df_problems.columns)))
    print("Processing complete. Check 'unique_maslulim.txt' and 'warnings.txt' for results.")
            
if __name__ == "__main__":
    main()