import pandas as pd

excel_file = pd.ExcelFile("PA-mizan.xlsx")
sheet_name = excel_file.sheet_names[2]

file_data = pd.read_excel("PA-mizan.xlsx", sheet_name)
df = pd.DataFrame(file_data)

#group_size = 150
#groups = [df[i:i + group_size].values.tolist() for i in range(0, len(df), group_size)]
#groups = [df.iloc[i:i + group_size].to_dict(orient='records') for i in range(0, len(df), group_size)]

df_sec= df[df["İşlem Grubu"] == "Seç"]
df_sinpas= df[(df["hesap kartno"].str.split('.').str[4] == "001087") & ( df["İşlem Grubu"] != "Seç")]
df = df[(df["hesap kartno"].str.split('.').str[4] != "001087") & ( df["İşlem Grubu"] != "Seç")]

group_size = 150
groups = [df.iloc[i:i + group_size].to_dict(orient='records') for i in range(0, len(df), group_size)]
groups_sec = [df_sec.iloc[i:i + group_size].to_dict(orient='records') for i in range(0, len(df_sec), group_size)]
groups_sinpas = [df_sinpas.iloc[i:i + group_size].to_dict(orient='records') for i in range(0, len(df_sinpas), group_size)]

all_groups = groups + groups_sec + groups_sinpas


print(len(all_groups))
