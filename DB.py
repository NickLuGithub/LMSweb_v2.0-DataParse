import pyodbc
import pandas as pd
import json

# 建立連線字串
conn_str = 'DRIVER={ODBC Driver 18 for SQL Server};SERVER=localhost;DATABASE=LMSContext;Trusted_Connection=yes;TrustServerCertificate=yes;'

# 建立連線
conn = pyodbc.connect(conn_str)

# 建立游標
cursor = conn.cursor()

# 執行SQL查詢
cursor.execute('SELECT * FROM EvaluationCoachings')

# 取得查詢結果
result = cursor.fetchall()

# 顯示查詢結果
sql = """SELECT DISTINCT 
            cou.CName
            ,[AUID]
            ,stu.GroupID
            ,g.GName
            ,[BUID]
            ,g2.GName
            ,[MissionID]
            ,[Evaluation]
            ,[Coaching]
        FROM EvaluationCoachings EC
        INNER JOIN Students stu on stu.StudentID = AUID
        INNER JOIN Students stu2 on stu2.StudentID = BUID
        INNER JOIN Groups g on stu.GroupID = g.GID
        INNER JOIN Groups g2 on stu2.GroupID = g2.GID
        INNER JOIN Missions mis on mis.MID = EC.MissionID
        INNER JOIN Courses cou on cou.CID = mis.CourseID
        WHERE cou.TeacherID = 'T005' and cou.TestType in ('2', '4', '5')
                and mis.MName in ('任務一:圖形製作','任務1:圖形製作', '任務一：圖型變化') 
        ORDER By cou.CName"""
df = pd.read_sql(sql, conn)

# PE01 PE02 PE03 PEA01 PEA02 PEC03 PEM07 PEIR08
PEQ = ['PE01', 'PE02', 'PE03', 'PEA01', 'PEA02', 'PEC03', 'PEM07', 'PEIR08']

for i in range(len(PEQ)):
    df[PEQ[i]] = df['Evaluation'].str.contains(PEQ[i])
    print(df['Evaluation'].str.contains(PEQ[i]))
    # 解析Evaluation欄位 json格式
    evaList = df['Evaluation'].str.split(",").tolist()
    print(evaList)
    df[PEQ[i]] = df['Evaluation'].str.contains(PEQ[i])

df.to_csv('EvaluationCoachings.csv', encoding='utf-8-sig')
# 關閉游標和連線
cursor.close()
conn.close()
