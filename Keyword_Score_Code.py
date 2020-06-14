from pandas import *
import csv

output = []

lifile=['Games.xlsx','Macro.xlsx','Rummy.xlsx']
li2=['Games', 'Macro', 'Rummy']
overall = {}
li2count = 0
for i in lifile:
    xls1 = ExcelFile(i)
    df1 = xls1.parse(xls1.sheet_names[0])
    lis1=[]

    for i in df1.values:
        lis1.append(i)
    dic={}
    for member in lis1:
        a=member[14][0:12]
        if a not in dic:
            game = li2[li2count]
            data= {"Mobile Premier League (MPL)": [0, 0, 0, 0]}

            dic[a]=data

        companies = dic[a]

        if member[10] == "Rewards":
            company_Name = "Mobile Premier League (MPL)"
            companies[company_Name][0] = companies[company_Name][0] + 1
        if member[11] == "Engagement":
            company_Name = "Mobile Premier League (MPL)"
            companies[company_Name][1] = companies[company_Name][1] + 1
        if member[12] == "Community based gaming":
            company_Name = "Mobile Premier League (MPL)"
            companies[company_Name][2] = companies[company_Name][2] + 1
        if member[13] == "Tournament":
            company_Name = "Mobile Premier League (MPL)"
            companies[company_Name][3] = companies[company_Name][3] + 1

    key=list(dic.keys())
    s={ }
    for k in range(len(key)):

        for j in range(0, 4):  # Adding one to all the count
            dic[key[k]]['Mobile Premier League (MPL)'].append(dic[key[k]]['Mobile Premier League (MPL)'][j]+1)

        for i in range(4,5 ):  # sum of scores on that date  dict
            s[(str(k) + str(i))] = (
                    dic[key[k]]['Mobile Premier League (MPL)'][i] + dic[key[k]]['Mobile Premier League (MPL)'][i+1]
                    + dic[key[k]]['Mobile Premier League (MPL)'][i+2] + dic[key[k]]['Mobile Premier League (MPL)'][i+3])

        for i in range(4,8): # (count / sum of scores on that date )*10 after calculating append to list
            dic[key[k]]['Mobile Premier League (MPL)'].append(round((dic[key[k]]['Mobile Premier League (MPL)'][i] / s[(str(k) + str(4))]) * 10,1))




    fields = ['start_date', 'end_date', 'brand', 'audience','keyword', 'score']



    engaged = ['Rewards', 'Engagement', 'Community based gaming', 'Tournament ']
    for key in dic:
        for key2 in dic[key]:
            scores = dic[key][key2]
            for j in range(len(engaged)):
                csvDict={}
                csvDict['start_date'] = key
                csvDict['end_date'] = key
                csvDict['brand'] = key2
                csvDict['audience'] = game
                csvDict['keyword'] = engaged[j]
                csvDict['score'] =scores[j+8]

                output.append(csvDict)

    overall[game] = dic

    li2count = li2count + 1

collectByDate = {}
for game in overall:
    for day in overall[game]:
        if day not in collectByDate:
            collectByDate[day] = {}
            for brand in overall[game][day]:
                collectByDate[day][brand] = []
        for brand in overall[game][day]:
            collectByDate[day][brand].append(overall[game][day][brand])


overallData = {}
for day in collectByDate:
    overallData[day] = {}
    for brand in collectByDate[day]:
        overallData[day][brand] = []
        overallData[day][brand].append("overall")
        overallData[day][brand].append(0)
        overallData[day][brand].append(0)
        overallData[day][brand].append(0)
        overallData[day][brand].append(0)
        for scores in collectByDate[day][brand]:
            for i in range(8, 12):
                overallData[day][brand][i - 7] = overallData[day][brand][i - 7] + scores[i]
        for i in range(1, 5):
            overallData[day][brand][i] = round((overallData[day][brand][i] / len(collectByDate[day][brand])), 1)


for day in overallData:
    for brand in overallData[day]:
        engaged = ['Rewards', 'Engagement', 'Community based gaming', 'Tournament ']
        for i in range(1,5):
            csvData = {}
            csvData['start_date'] = day
            csvData['end_date'] = day
            csvData['brand'] = brand
            csvData['audience'] = overallData[day][brand][0]
            csvData['keyword'] = engaged[i-1]
            csvData['score'] = overallData[day][brand][i]

            output.append(csvData)


filename2 = "Output_Keyword_Score.csv"

with open(filename2, 'w', newline='') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=fields)
    writer.writeheader()
    writer.writerows(output)

