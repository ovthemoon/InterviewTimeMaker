import pandas as pd

df = pd.DataFrame(columns=['A', 'B', 'C'])
df = df.append({'A': 1, 'B': 2, 'C': 3}, ignore_index=True)
print(df)
