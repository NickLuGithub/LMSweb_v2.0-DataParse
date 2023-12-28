import pyodbc
import pandas as pd
import json

# 建立連線字串
conn_str = 'DRIVER={ODBC Driver 18 for SQL Server};SERVER=localhost;DATABASE=LMSContext;Trusted_Connection=yes;TrustServerCertificate=yes;'

# 建立連線
conn = pyodbc.connect(conn_str)

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

# # 將Evaluation欄位前後加上{}
# df['Evaluation'] = df['Evaluation'].apply(lambda x: json.dumps(x))

# # 拆分Evaluation欄位成PE01 PE02 PE03 PEA01 PEA02 PEC03 PEM07 PEIR08欄位
# df = pd.concat([df.drop(['Evaluation'], axis=1), df['Evaluation'].apply(lambda x: pd.Series(json.loads(x)))], axis=1)

# 將Evaluation欄位，用,:分割成多個欄位，產生出來的欄位只保留:右手邊的值，並且刪除原本的Evaluation欄位
df = pd.concat([df.drop(['Evaluation'], axis=1), df['Evaluation'].str.split(':', expand=True)], axis=1)
# 將Coaching欄位 用,:分割成多個欄位 並且刪除原本的Coaching欄位
df = pd.concat([df.drop(['Coaching'], axis=1), df['Coaching'].str.split(':', expand=True)], axis=1)

df.to_csv('EvaluationCoachings.csv', encoding='utf-8-sig')