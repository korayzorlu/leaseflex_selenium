import pandas as pd

# ❶ Kaynak dosya
source_path = "PA-mizan.xlsx"          # <-- dosyanızı buraya yazın
sheet_name  = 2                      # veya "Sayfa1" vb.

# ❷ Kaç satırda bir bölünecek?
chunk_size  = 150

# ❸ Veriyi yükle
df = pd.read_excel(source_path, sheet_name=sheet_name)

# ❹ Kaç parça çıkacağını hesapla
total_rows = len(df)
num_chunks = (total_rows + chunk_size - 1) // chunk_size   # tam bölünmezse son parçayı da say

# ❺ Her parçayı yaz
for i in range(num_chunks):
    start = i * chunk_size
    end   = start + chunk_size
    chunk = df.iloc[start:end]

    out_path = f"mizan/mizan_{i+1:03}.xlsx"     # parca_001.xlsx, parca_002.xlsx …
    chunk.to_excel(out_path, index=False)
    print(f"{out_path} kaydedildi ({len(chunk)} satır)")