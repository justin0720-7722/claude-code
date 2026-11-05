# -*- coding: utf-8 -*-
"""
HW2 校園社團活動管理系統 — 資料表 / 欄位規格（單一來源，依 HW1_TableSchemaSketch_41411017E 設計）
Excel 工作檔、匯入檔、Access 資料庫都由這份規格產生。

公式樣板佔位符：
  {F:欄位}      同一列另一個欄位（或輔助欄）的儲存格，例如 {F:Grade} -> D5
  {C:欄位}      該欄位的欄字母，例如 {C:ActivityID} -> B
  {R:表.欄位}   另一張表整欄的絕對參照，例如 {R:StudentTable.StudentID} -> StudentTable!$A:$A
  {P}           上一列的列號（用於「本列以上」的累計範圍）
"""
import datetime as dt

STUDENT_ID = "41411017E"
BASE_DATE = dt.datetime(2026, 11, 5)          # 系統基準日（「今天」），固定下來讓資料可重現
SEMESTER = "115 學年度第 1 學期"

# ---------------------------------------------------------------- 參數清單（皆定義為 Excel 名稱）
DEPTS = ["資訊管理學系", "資訊工程學系", "企業管理學系", "國際企業學系", "英語學系", "數學系",
         "物理學系", "化學系", "生命科學系", "美術學系", "音樂學系", "體育學系", "教育學系",
         "心理與諮商學系", "地理學系", "歷史學系"]
CLUBS = [  # (社團名稱, 社團類別)
    ("程式設計研究社", "學術"), ("資訊安全研究社", "學術"), ("數學研究社", "學術"), ("天文社", "學術"),
    ("辯論社", "學術"), ("經濟研究社", "學術"), ("日語研究社", "學術"),
    ("攝影社", "藝文"), ("吉他社", "藝文"), ("合唱團", "藝文"), ("管樂社", "藝文"), ("話劇社", "藝文"),
    ("書法社", "藝文"), ("動漫社", "藝文"), ("熱舞社", "藝文"),
    ("籃球社", "體育"), ("排球社", "體育"), ("羽球社", "體育"), ("桌球社", "體育"), ("游泳社", "體育"),
    ("登山社", "體育"), ("跆拳道社", "體育"),
    ("慈幼社", "服務"), ("春暉社", "服務"), ("環保志工社", "服務"), ("國際志工社", "服務"),
    ("康輔社", "康樂"), ("桌遊社", "康樂"), ("魔術社", "康樂"), ("南友會", "聯誼"),
]
CLUB_TYPES = ["學術", "藝文", "體育", "服務", "康樂", "聯誼"]
ACT_STATUS = ["報名中", "已額滿", "已結束", "取消"]
ATTEND = ["未簽到", "已出席", "缺席"]
PAY = ["未繳", "已繳", "免繳"]

LISTS = {
    "姓氏": list("陳林黃張李王吳劉蔡楊許鄭謝郭洪曾邱廖賴周徐蘇葉莊呂江何蕭羅高潘簡朱鍾游彭詹胡施沈"),
    "名字": ["志明", "家豪", "俊傑", "建宏", "承恩", "冠宇", "宇翔", "柏翰", "彥廷", "品睿",
           "子軒", "宥廷", "哲瑋", "承翰", "宗翰", "育誠", "立偉", "信宏", "睿恩", "致遠",
           "怡君", "雅婷", "佳穎", "詩涵", "欣怡", "宜蓁", "思妤", "子晴", "品妍", "雨潔",
           "芷瑄", "郁婷", "佩珊", "家瑜", "筱涵", "語彤", "依晨", "心怡", "羽彤", "亭妤"],
    "系所清單": DEPTS,
    "系所代碼": [f"{i:02d}" for i in range(1, len(DEPTS) + 1)],
    "社團清單": [c[0] for c in CLUBS],
    "社團類別對照": [c[1] for c in CLUBS],
    "社團類別清單": CLUB_TYPES,
    "簡介片語一": ["以推廣興趣交流為宗旨", "致力培養專業技能", "歡迎零基礎同學加入", "每週固定舉辦社課",
              "由學長姐帶領學習", "重視團隊合作精神", "提供跨系交流平台", "結合理論與實作"],
    "簡介片語二": ["每學期舉辦成果發表", "定期舉辦校外交流", "積極參與全國大專競賽", "結合服務學習課程",
              "打造溫暖的社團氛圍", "提供器材借用服務", "與他校社團合辦活動", "培養未來社團幹部"],
    "活動類型": ["期初社員大會", "新生迎新茶會", "入門教學課程", "進階工作坊", "專題講座", "校外參訪",
              "成果發表會", "聯誼晚會", "志工服務", "期末聚餐", "校內友誼賽", "讀書會"],
    "地點清單": ["綜合教學大樓A101", "綜合教學大樓B203", "學生活動中心多功能廳", "學生活動中心社辦",
              "體育館", "游泳館", "圖書館國際會議廳", "誠正大樓演講廳", "戶外操場", "線上會議室"],
    "活動狀態清單": ACT_STATUS,
    "出席狀態清單": ATTEND,
    "繳費狀態清單": PAY,
    "備註清單": ["第一次參加活動", "需要素食餐點", "與同學一起報名", "已加入社團LINE群組",
              "需借用器材", "可能晚到 15 分鐘", "需要停車位", "請開立收據"],
    "基準日": [BASE_DATE],
}


def pick(listname):
    return f"INDEX({listname},RANDBETWEEN(1,ROWS({listname})))"


def fk(table, field):
    """從父表的主鍵欄隨機挑一筆（範圍隨父表筆數自動調整）。"""
    return f"INDEX({{R:{table}.{field}}},RANDBETWEEN(2,COUNTA({{R:{table}.{field}}})))"


def vlist(items):
    return 'In (' + ",".join(f'"{x}"' for x in items) + ')'


def F(name, caption, typ, size=None, *, req=False, key="", index=False, rule="", rtext="",
      fmt="", desc="", formula="", lookup=None):
    """typ: TEXT / MEMO / BYTE / INT / LONG / AUTONUMBER / MONEY / DATE / DATETIME
    key: PK / UQ / FK:表.欄位；lookup: ("values", [...]) 或 ("query", SQL, 欄數, 欄寬)"""
    return dict(name=name, caption=caption, type=typ, size=size, req=req, key=key, index=index,
                rule=rule, rtext=rtext, fmt=fmt, desc=desc, formula=formula, lookup=lookup)


# ---------------------------------------------------------------- 資料表規格
TABLES = [
    {
        "name": "StudentTable", "zh": "學生資料", "rows": 60, "desc": "記錄學生基本資料",
        "fields": [
            F("StudentID", "學號", "TEXT", 10, req=True, key="PK",
              rule='Like "4#######[A-Z]"', rtext="學號格式：4 + 入學年2碼 + 系所2碼 + 序號3碼 + 英文字母",
              desc="主鍵；入學年與年級一致（115 學年度 1 年級 = 15）",
              formula='="4"&TEXT(116-{F:Grade}-100,"00")&INDEX(系所代碼,MATCH({F:DeptName},系所清單,0))&TEXT(ROW()-1,"000")&"E"'),
            F("StudentName", "姓名", "TEXT", 30, req=True, desc="學生姓名",
              formula="=" + pick("姓氏") + "&" + pick("名字")),
            F("DeptName", "系所", "TEXT", 30, index=True, rule=vlist(DEPTS), rtext="請選擇校內系所名稱",
              desc="限制為校內系所名稱（查閱清單）", formula="=" + pick("系所清單"),
              lookup=("values", DEPTS)),
            F("Grade", "年級", "BYTE", req=False, rule="Between 1 And 4", rtext="年級只能填 1～4",
              desc="限制 1～4", formula="=RANDBETWEEN(1,4)"),
            F("Email", "電子郵件", "TEXT", 100, key="UQ", rule='Is Null Or Like "*@*.*"', rtext="Email 格式錯誤",
              desc="由學號產生，設定唯一避免重複", formula='=LOWER({F:StudentID})&"@example.edu.tw"'),
            F("Phone", "聯絡電話", "TEXT", 15, rule='Is Null Or Like "09########"', rtext="手機須為 09 開頭 10 碼",
              desc="可空值（約 10% 未填，用來測試空值）",
              formula='=IF(RANDBETWEEN(1,10)=1,"","09"&TEXT(RANDBETWEEN(0,99999999),"00000000"))'),
        ],
        "helpers": [("輔助_社長亂數", "=IF({F:Grade}>=2,RAND(),-1)")],
    },
    {
        "name": "ClubTable", "zh": "社團資料", "rows": 30, "desc": "記錄社團基本資料",
        "fields": [
            F("ClubID", "社團編號", "AUTONUMBER", req=True, key="PK", desc="主鍵；自動編號（1001 起）",
              formula="=1000+ROW()-1"),
            F("ClubName", "社團名稱", "TEXT", 50, req=True, key="UQ", desc="社團名稱不可重複",
              formula="=INDEX(社團清單,ROW()-1)"),
            F("ClubType", "社團類別", "TEXT", 10, req=True, index=True, rule=vlist(CLUB_TYPES),
              rtext="社團類別不在清單內", desc="查閱：學術、藝文、體育、服務、康樂、聯誼",
              formula="=INDEX(社團類別對照,ROW()-1)", lookup=("values", CLUB_TYPES)),
            F("PresidentStudentID", "社長學號", "TEXT", 10, req=True, key="FK:StudentTable.StudentID",
              desc="外鍵；社長須為 2 年級以上，且一人只擔任一個社團社長",
              formula="=INDEX({R:StudentTable.StudentID},MATCH(LARGE({R:StudentTable.輔助_社長亂數},ROW()-1),{R:StudentTable.輔助_社長亂數},0))",
              lookup=("query", "SELECT StudentID, StudentName FROM StudentTable ORDER BY StudentID;", 2, "1440;1440")),
            F("FoundedYear", "創立年份", "INT", rule="Between 1946 And Year(Date())", rtext="創立年份不合理",
              desc="社團成立年份（1960～2024）", formula="=RANDBETWEEN(1960,2024)"),
            F("Description", "社團簡介", "MEMO", desc="長文字；記錄社團宗旨與活動特色",
              formula='={F:ClubName}&"："&' + pick("簡介片語一") + '&"，"&' + pick("簡介片語二") + '&"。"'),
        ],
        "uniques": [["PresidentStudentID"]],
    },
    {
        "name": "ActivityTable", "zh": "活動資料", "rows": 60, "desc": "記錄社團舉辦的活動",
        "fields": [
            F("ActivityID", "活動編號", "AUTONUMBER", req=True, key="PK", desc="主鍵；自動編號（2001 起）",
              formula="=2000+ROW()-1"),
            F("ClubID", "社團編號", "LONG", req=True, key="FK:ClubTable.ClubID", desc="外鍵；連結 ClubTable",
              formula="=" + fk("ClubTable", "ClubID"),
              lookup=("query", "SELECT ClubID, ClubName FROM ClubTable ORDER BY ClubID;", 2, "1008;1800")),
            F("ActivityName", "活動名稱", "TEXT", 100, req=True, index=True, desc="社團名稱＋活動類型",
              formula="=INDEX({R:ClubTable.ClubName},MATCH({F:ClubID},{R:ClubTable.ClubID},0))&" + pick("活動類型")),
            F("ActivityDate", "活動日期", "DATE", req=True, index=True,
              rule="Between #9/1/2026# And #1/31/2027#", rtext=f"活動日期須在{SEMESTER}內",
              fmt="Short Date", desc=f"{SEMESTER}（2026/09/14～2026/12/20）",
              formula="=RANDBETWEEN(DATE(2026,9,14),DATE(2026,12,20))"),
            F("Location", "活動地點", "TEXT", 100, desc="活動舉辦地點", formula="=" + pick("地點清單")),
            F("MaxParticipants", "最大人數", "INT", req=True, rule=">0", rtext="最大人數必須大於 0",
              desc="大於 0 的整數；不可小於實際報名人數（報名 ≥3 人的活動約 1/4 剛好額滿）",
              formula="=IF(AND(RANDBETWEEN(1,4)=1,{F:輔助_報名人數}>=3),{F:輔助_報名人數},CEILING({F:輔助_報名人數}+RANDBETWEEN(1,20),5))"),
            F("Fee", "活動費用", "MONEY", req=True, rule=">=0", rtext="費用不可為負數", fmt="Currency",
              desc="無費用為 0（約 4 成免費），其餘 50～500 元",
              formula="=IF(RANDBETWEEN(1,10)<=4,0,RANDBETWEEN(1,10)*50)"),
            F("ActivityStatus", "活動狀態", "TEXT", 10, req=True, index=True, rule=vlist(ACT_STATUS),
              rtext="活動狀態不在清單內", desc="依日期與報名人數推導；5% 取消",
              formula='=IF(RANDBETWEEN(1,20)=1,"取消",IF({F:ActivityDate}<基準日,"已結束",IF({F:輔助_報名人數}>={F:MaxParticipants},"已額滿","報名中")))',
              lookup=("values", ACT_STATUS)),
        ],
        "helpers": [
            ("輔助_報名人數", "=COUNTIF({R:RegistrationTable.ActivityID},{F:ActivityID})"),
            ("輔助_起始學生", "=RANDBETWEEN(0,COUNTA({R:StudentTable.StudentID})-2)"),
        ],
    },
    {
        "name": "RegistrationTable", "zh": "活動報名", "rows": 240, "desc": "記錄學生參加活動的報名關係",
        "fields": [
            F("RegistrationID", "報名編號", "AUTONUMBER", req=True, key="PK", desc="主鍵；自動編號（3001 起）",
              formula="=3000+ROW()-1"),
            F("ActivityID", "活動編號", "LONG", req=True, key="FK:ActivityTable.ActivityID",
              desc="外鍵；前 60 筆依序對應每場活動（每場至少 1 人），其餘隨機",
              formula="=IF(ROW()<=COUNTA({R:ActivityTable.ActivityID}),INDEX({R:ActivityTable.ActivityID},ROW()),"
                      + fk("ActivityTable", "ActivityID") + ")",
              lookup=("query", "SELECT ActivityID, ActivityName FROM ActivityTable ORDER BY ActivityID;", 2, "1008;2880")),
            F("StudentID", "學號", "TEXT", 10, req=True, key="FK:StudentTable.StudentID",
              desc="外鍵；同一活動依「起始學生 + 第幾位 × 7」取學生，保證不重複報名",
              formula="=INDEX({R:StudentTable.StudentID},2+MOD({F:輔助_起始學生}+COUNTIF({C:ActivityID}$1:{C:ActivityID}{P},{F:ActivityID})*7,COUNTA({R:StudentTable.StudentID})-1))",
              lookup=("query", "SELECT StudentID, StudentName FROM StudentTable ORDER BY StudentID;", 2, "1440;1440")),
            F("RegisterTime", "報名時間", "DATETIME", req=True, rule=">=#9/1/2026#", rtext="報名時間不可早於開學日",
              fmt="yyyy/mm/dd hh:nn", desc="活動前 45 天內、活動前一天以前、且不晚於基準日；時間 08:00～23:59",
              formula="=RANDBETWEEN(MAX({F:輔助_活動日期}-45,DATE(2026,9,1)),MIN({F:輔助_活動日期}-1,基準日))+RANDBETWEEN(480,1439)/1440"),
            F("AttendanceStatus", "出席狀態", "TEXT", 5, req=True, index=True, rule=vlist(ATTEND),
              rtext="出席狀態不在清單內", desc="未舉辦或取消的活動一律未簽到；已結束 85% 出席",
              formula='=IF(OR({F:輔助_活動狀態}="取消",{F:輔助_活動日期}>=基準日),"未簽到",IF(RANDBETWEEN(1,100)<=85,"已出席","缺席"))',
              lookup=("values", ATTEND)),
            F("PaymentStatus", "繳費狀態", "TEXT", 5, req=True, index=True, rule=vlist(PAY),
              rtext="繳費狀態不在清單內", desc="免費活動一律免繳；有出席者一律已繳",
              formula='=IF({F:輔助_活動費用}=0,"免繳",IF({F:AttendanceStatus}="已出席","已繳",IF(RANDBETWEEN(1,10)<=7,"已繳","未繳")))',
              lookup=("values", PAY)),
            F("Note", "備註", "TEXT", 200, desc="可空值；約 25% 有備註",
              formula='=IF(RANDBETWEEN(1,4)=1,' + pick("備註清單") + ',"")'),
        ],
        "helpers": [
            ("輔助_起始學生", "=INDEX({R:ActivityTable.輔助_起始學生},MATCH({F:ActivityID},{R:ActivityTable.ActivityID},0))"),
            ("輔助_活動日期", "=INDEX({R:ActivityTable.ActivityDate},MATCH({F:ActivityID},{R:ActivityTable.ActivityID},0))"),
            ("輔助_活動費用", "=INDEX({R:ActivityTable.Fee},MATCH({F:ActivityID},{R:ActivityTable.ActivityID},0))"),
            ("輔助_活動狀態", "=INDEX({R:ActivityTable.ActivityStatus},MATCH({F:ActivityID},{R:ActivityTable.ActivityID},0))"),
        ],
        "uniques": [["ActivityID", "StudentID"]],
    },
]

DESIGN_CHANGES = [
    ("學號加上格式驗證，且入學年與年級連動",
     "HW1 只設定 短文字(10)。HW2 加上 Like \"4#######[A-Z]\"，並讓學號中的入學年 = 116 − 年級（115 學年度 1 年級 → 15），避免「學號是 112 入學卻填 1 年級」。"),
    ("系所改成「查閱 + 驗證規則」",
     "HW1 寫「限制為校內系所名稱」，實作成值清單下拉選單，再加 In(...) 驗證規則，打錯字的系所名稱進不來。"),
    ("Email 由學號自動組成",
     "學號已唯一，Email = 學號小寫 + @example.edu.tw，自然不會重複，符合 HW1 的唯一限制。"),
    ("聯絡電話加上格式驗證，並保留空值",
     "HW1 設定可空值，所以驗證規則寫成 Is Null Or Like \"09########\"；假資料刻意讓約 10% 空白，測試空值能否正常匯入。"),
    ("社長學號加上唯一索引、限 2 年級以上",
     "HW1 只設外鍵。實務上一個學生不會同時當兩個社團社長，新生也不會當社長，所以加唯一索引，假資料也只從 2 年級以上抽。"),
    ("創立年份：數字(4) → 整數 + 範圍驗證",
     "Access 的「數字」要選欄位大小，4 位數年份用「整數」即可，並限制 1946（建校）～今年。"),
    ("活動費用：數字 → 貨幣",
     "金額用「貨幣」型態避免浮點誤差，並限制 >= 0。"),
    ("最大人數不可小於實際報名人數",
     "HW1 只限制 > 0。HW2 讓最大人數依報名人數產生，並在「檢核」頁確認沒有超收。"),
    ("活動狀態改由資料推導",
     "已過基準日 → 已結束；報名人數 = 最大人數 → 已額滿；其餘 → 報名中；另有 5% 取消。避免出現「已結束但日期在未來」。"),
    ("活動日期限定在本學期",
     f"所有活動都在 {SEMESTER}，這樣 1 年級新生（2026/9 入學）也能合理報名，不會出現「還沒入學就報名」。"),
    ("報名時間必須早於活動日期",
     "報名時間 = 活動前 45 天內的隨機日期 + 08:00～23:59 的隨機時間，且不晚於基準日（不能在未來報名）。"),
    ("出席 / 繳費狀態與活動連動",
     "未舉辦或取消的活動只能是「未簽到」；免費活動一律「免繳」；有出席的一律「已繳」。"),
    ("(活動編號, 學號) 複合唯一用公式保證",
     "HW1 要求同一學生不可重複報名同一活動。隨機抽學生一定會撞，所以同一活動的第 k 位報名者取「起始學生 + k×7」（7 與學生人數 60 互質），保證不重複，並在 Access 建立複合唯一索引。"),
    ("自動編號從 1001 / 2001 / 3001 開始",
     "沿用 HW1 的填寫範例（社團 1001、活動 2001、報名 3001），一看編號就知道是哪張表。"),
]
