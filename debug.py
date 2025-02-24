import pandas as pd
filepath = ''
debug_key1 = "26713B6F4CCFE7E287FAFB97888F4841"
debug_key2 = "test1"
df = pd.read_excel(filepath, keep_default_na=False)
filtered_df = df[df['Key'] == debug_key2]
print(filtered_df['MsgStr'])
