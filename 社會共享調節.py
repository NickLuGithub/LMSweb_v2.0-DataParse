from sqlalchemy import create_engine, text
import csv
import io

# 設定資料庫連線資訊
host = 'localhost'
database = 'LMSContext'
driver = 'ODBC Driver 18 for SQL Server'

# 建立連線
engine = create_engine(f'mssql+pyodbc://{host}/{database}?driver={driver}&TrustServerCertificate=yes')

query = text("""SELECT c.CName,s.StudentName,g.GName,m.MName, q.QuestionID ,q.QContent, a.Acontent,ep.Name
FROM Answers a
INNER JOIN Questions q on q.QuestionID = a.QuestionID
INNER JOIN ExperimentalProcedures ep on ep.EProcedureID = q.EProcedureID
INNER JOIN Missions m on m.MID = a.MissionID
INNER JOIN Executions ex on ex.MissionID = m.MID
INNER JOIN Groups g on g.GID = ex.GroupID
INNER JOIN Students s on s.StudentID = a.UserID and s.GroupID = g.GID
INNER JOIN Courses c on c.CID = s.CourseID and g.CourseID = c.CID
WHERE s.isLeader = 1 and m.MName in ('任務一:圖形製作', '任務1:圖形製作', '任務一：圖型變化') 
ORDER BY c.CName, g.GID, ep.Name""")

# 現在可以使用 engine 來執行 SQL 查詢等操作
with engine.connect() as connection:
    result = connection.execute(query)
    # 使用 StringIO 作為一個暫存的輸出文件
    output = io.StringIO()

    # 創建一個 CSV 寫入器
    writer = csv.writer(output)

    # 寫入查詢結果的標題 (即列名)
    writer.writerow(result.keys())

    # 寫入查詢結果的每一行
    writer.writerows(result.fetchall())

    # 獲取 CSV 字串並轉換為 UTF-8-SIG 編碼
    csv_content = output.getvalue().encode('utf-8-sig')

# 處理或保存 csv_content
# 例如，將其寫入文件：
with open('output.csv', 'wb') as file:
    file.write(csv_content)

# pd 讀取 csv 檔案
import pandas as pd
df = pd.read_csv('output.csv', encoding='utf-8-sig')
print(df)

# 同一位同學的 QuestionID 會有重複的，因為這一個 QuestionID 有多個 Answer，所以要把重複的 QuestionID 去除，只留下一個，並且把重複的 Answer 內容合併成一個欄位
df = df.groupby(['QuestionID', 'CName', 'GName'])['Acontent'].apply(' '.join).reset_index()
# 刪除 GName 欄位中的 '第' 和 '組'
df['GName'] = df['GName'].str.replace('第', '')
df['GName'] = df['GName'].str.replace('組', '')
df['GName'] = df['GName'].str.zfill(2)

# 宣告一轉換表，紀錄中文數字對應到的阿拉伯數字
change = {'一': '1', '二': '2', '三': '3', '四': '4', '五': '5', '六': '6', '七': '7', '八': '8', '九': '9', '十': '10'}

# 將 CName 欄位中的中文數字轉換成阿拉伯數字
for i in range(len(df)):
    for key, value in change.items():
        df['CName'][i] = df['CName'][i].replace(key, value)

# 將 CName 欄位中的 '年' 與 '班'
df['CName'] = df['CName'].str.replace('年', '0')
df['CName'] = df['CName'].str.replace('班', '')


order = ['PGS01B', 'PGS02B', 'PGS03B', 'TM01B', 'CM01B', 'TE01B', 'TR01B', 'TR02B', 'TR03B', 'TR04B']
df['QuestionID'] = pd.Categorical(df['QuestionID'], categories=order, ordered=True)
df = df.sort_values('QuestionID')
print(df)
df = df.sort_values(by=['CName', 'GName'])
print(df)

# 輸出成 csv 檔案
df.to_csv('output2.csv', encoding='utf-8-sig', index=False)