import pandas as pd
df = pd.read_excel(r"D:\桌面\700人大名单简要信息汇总-学校.xlsx")
name = df.iloc[0]["申请人"]
company = df.iloc[0]["单位"]