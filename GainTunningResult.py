import csv
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from matplotlib import pyplot as plt

# Ver0307
RT_Score = []
# Response time score calculation
def CalScore_RT(raw_data) :
    raw_data[1] = int(raw_data[1])
    raw_data[2] = int(raw_data[2])
    score_val = 0

    if raw_data[1] == 0 : #PR
        if raw_data[2] > 250 :
            score_val = 100-abs(raw_data[2]-250)
        else :
            score_val = 100
    elif raw_data[1] == 1: #PN
        if raw_data[2] > 300:
            score_val = 100-abs(raw_data[2] - 300)
        else:
            score_val = 100
    elif raw_data[1] == 2: #PD
        if raw_data[2] > 450:
            score_val = 100-abs(raw_data[2] - 450)
        else:
            score_val = 100
    elif raw_data[1] == 3: #RP
        if raw_data[2] > 1000:
            score_val = 100-abs(raw_data[2] - 1000)
        else:
            score_val = 100
    elif raw_data[1] == 4: #RN
        if raw_data[2] > 200:
            score_val = 100-abs(raw_data[2] - 200)
        else:
            score_val = 100
    elif raw_data[1] == 5:  # RD
        if raw_data[2] > 350:
            score_val = 100-abs(raw_data[2] - 350)
        else:
            score_val = 100
    elif raw_data[1] == 6:  # NP
        if raw_data[2] > 300:
            score_val = 100-abs(raw_data[2] - 300)
        else:
            score_val = 100
    elif raw_data[1] == 7:  # NR
        if raw_data[2] > 200:
            score_val = 100-abs(raw_data[2] - 200)
        else:
            score_val = 100
    elif raw_data[1] == 8:  # ND
        if raw_data[2] > 200:
            score_val = 100-abs(raw_data[2] - 200)
        else:
            score_val = 100
    elif raw_data[1] == 9:  # DP
        if raw_data[2] > 1000:
            score_val = 100-abs(raw_data[2] - 1000)
        else:
            score_val = 100
    elif raw_data[1] == 10:  # DR
        if raw_data[2] > 350:
            score_val = 100-abs(raw_data[2] - 350)
        else:
            score_val = 100
    elif raw_data[1] == 11:  # DN
        if raw_data[2] > 200:
            score_val = 100-abs(raw_data[2] - 200)
        else:
            score_val = 100

    list = [int(raw_data[0]), score_val]
    RT_Score.append(list)


SA_Score = []
# Stop Accuracy score calculation
def CalScore_SA(raw_data) :
    score_val = 0
    raw_data[4] = float(raw_data[4])

    if raw_data[4] == 0 :
        score_val = 100
    else :
        score_val = 100 - abs(raw_data[4]) * 10

    list = [int(raw_data[0]), score_val]
    SA_Score.append(list)

OS_Score = []
# Overshoot score calculation
def CalScore_OS(raw_data) :
    raw_data[6] = float(raw_data[6])

    if raw_data[6] == 0 :
        score_val = 100
    else :
        score_val = 100 - abs(raw_data[6]) * 10

    list = [int(raw_data[0]), score_val]
    OS_Score.append(list)


def FindBestGainSet(sum_score):
    max = 0
    for idx in sum_score :
        val = int(idx[1])
        if max < val :
            max = val
            gainset = int(idx[0])
        else:
            continue
    print("[%d] gainset (score : %d) is the Best Gain Set!" %(gainset,max))

# 엑셀 데이터 저장
WB = Workbook()
WS = WB.active
WS.title = "SQ_Auto_Gaintunning"

Description = ["Gain Set", "항목", "PR", "PN", "PD", "RP", "RN", "RD", "NP", "NR", "ND", "DP", "DR", "DN", "총점수"]

for col in range(0, len(Description)):
    WS.cell(0+1, col+1).value = Description[col]

sum_score_RT = 0
sum_score_SA = 0
sum_score_OS = 0
log_count = 0
sum_score = []
row = 0
# read excel file
with open('./TEST_RESULT.csv', 'r') as file :
    log = csv.reader(file)

    for raw_data in log :
        if log_count >= 1 :
            CalScore_RT(raw_data)
            CalScore_SA(raw_data)
            CalScore_OS(raw_data)

        log_count = log_count + 1

    start = 0
    end = len(RT_Score)
    div = 12

    for idx in range(start, end, div):
        resRT = RT_Score[start:start+div]
        resSA = SA_Score[start:start+div]
        resOS = OS_Score[start:start+div]

        # gain set 별 제어 성능 점수 계산
        if resRT != [] :
            for i in range(len(resRT)) :
                gainset = resRT[i][0]
                val = resRT[i][1]
                sum_score_RT += val
            print("[%d] Gainset Response Time Score : %d" % (gainset,sum_score_RT))

        if resSA != [] :
            for i in range(len(resSA)) :
                gainset = resSA[i][0]
                val = resSA[i][1]
                sum_score_SA += val
            print("[%d] Gainset Stop Accuracy Score : %d" % (gainset, sum_score_SA))

        if resOS != []:
            for i in range(len(resOS)):
                gainset = resOS[i][0]
                val = resOS[i][1]
                sum_score_OS += val
            print("[%d] Gainset Overshoot Score : %d" % (gainset, sum_score_OS))

        # Excel 에 data 저장
        col = 0
        WS.append([resRT[col][0], "응답시간", resRT[col][1], resRT[col + 1][1], resRT[col + 2][1], resRT[col + 3][1],
                   resRT[col + 4][1], resRT[col + 5][1], resRT[col + 6][1], resRT[col + 7][1], resRT[col + 8][1],
                   resRT[col + 9][1], resRT[col + 10][1], resRT[col + 11][1],sum_score_RT])
        WS.append([resSA[col][0], "제어정밀도", resSA[col][1], resSA[col + 1][1], resSA[col + 2][1], resSA[col + 3][1],
                   resSA[col + 4][1], resSA[col + 5][1], resSA[col + 6][1], resSA[col + 7][1], resSA[col + 8][1],
                   resSA[col + 9][1], resSA[col + 10][1], resSA[col + 11][1],sum_score_SA])
        WS.append([resOS[col][0], "오버슈트", resOS[col][1], resOS[col + 1][1], resOS[col + 2][1], resOS[col + 3][1],
                   resOS[col + 4][1], resOS[col + 5][1], resOS[col + 6][1], resOS[col + 7][1], resOS[col + 8][1],
                   resOS[col + 9][1], resOS[col + 10][1], resOS[col + 11][1],sum_score_OS])


        list = [gainset, int(sum_score_OS + sum_score_RT + sum_score_SA)]
        sum_score.append(list)
        sum_score_RT = 0
        sum_score_SA = 0
        sum_score_OS = 0
        start = start + div

    FindBestGainSet(sum_score)


WB.save("SQ GainTunning Test Result.xlsx")
