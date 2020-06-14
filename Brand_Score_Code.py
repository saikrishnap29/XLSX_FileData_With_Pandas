from pandas import *
import csv

output = []

lifile=['Macro.xlsx','Games.xlsx','Rummy.xlsx']
li2=['Macro','Games','Rummy']
overall = {}
li2count = 0
for i in lifile:
    xls1 = ExcelFile(i)
    df1 = xls1.parse(xls1.sheet_names[0])
    lis1=[]
    for i in df1.values:
        lis1.append(i)

    dic={}
    game = li2[li2count]
    for member in lis1:
        a=member[14][0:12]
        if a not in dic:
            data= {"Mobile Premier League (MPL)": [game, 0, 0, 0], "Adda52": [game, 0, 0, 0],
                  "Dream 11": [game, 0, 0, 0], "Rummy Circle": [game, 0, 0, 0]}

            dic[a]=data

        companies = dic[a]
        brand = companies[member[0]]
        brand[1] = brand[1] + 1

        if member[2] == "Option1":
            company_Name = "Mobile Premier League (MPL)"
            companies[company_Name][2] = companies[company_Name][2] + 1
        if member[3] == "Option2":
            company_Name = "Adda52"
            companies[company_Name][2] = companies[company_Name][2] + 1
        if member[4] == "Option3":
            company_Name = "Dream 11"
            companies[company_Name][2] = companies[company_Name][2] + 1
        if member[5] == "Option4":
            company_Name = "Rummy Circle"
            companies[company_Name][2] = companies[company_Name][2] + 1

        for i in range(6, 10):
            if notna(member[i]):
               companies[member[i]][3] = companies[member[i]][3] + 1

    key=list(dic.keys())
    s={}        # sum of scores on that date
    for k in range(len(key)):
        for j in range(1, 4): # Adding one to all the count
            dic[key[k]]['Mobile Premier League (MPL)'].append(dic[key[k]]['Mobile Premier League (MPL)'][j]+1)
            dic[key[k]]['Adda52'].append(dic[key[k]]['Adda52'][j]+1)
            dic[key[k]]['Dream 11'].append(dic[key[k]]['Dream 11'][j]+1)
            dic[key[k]]['Rummy Circle'].append(dic[key[k]]['Rummy Circle'][j]+1)
        for i in range(4,7):  # sum of scores on that date  dict
            s[(str(k)+str(i))]=(dic[key[k]]['Mobile Premier League (MPL)'][i]+dic[key[k]]['Adda52'][i]+dic[key[k]]['Dream 11'][i]+dic[key[k]]['Rummy Circle'][i])
        for i in range(4,7): # (count / sum of scores on that date )*10 after calculating append to list
            dic[key[k]]['Mobile Premier League (MPL)'].append(round(((dic[key[k]]['Mobile Premier League (MPL)'][i] / s[(str(k) + str(i))]) * 10),1))
            dic[key[k]]['Adda52'].append(round(((dic[key[k]]['Adda52'][i] / s[(str(k) + str(i))]) * 10),1))
            dic[key[k]]['Dream 11'].append(round(((dic[key[k]]['Dream 11'][i] / s[(str(k) + str(i))]) * 10),1))
            dic[key[k]]['Rummy Circle'].append(round(((dic[key[k]]['Rummy Circle'][i] / s[(str(k) + str(i))]) * 10),1))

        for i in range(7,8):  #  (FINAL-unaided_brand_awareness_score*5+FINAL-aided_brand_awareness_score*3+FINAL-brand_consideration*2)/10

            dic[key[k]]['Mobile Premier League (MPL)'].append(round((((dic[key[k]]['Mobile Premier League (MPL)'][i]) * 5) +
                ((dic[key[k]]['Mobile Premier League (MPL)'][i + 1]) * 3) + ((dic[key[k]]['Mobile Premier League (MPL)'][i + 2]) * 2)) / 10,1))

            dic[key[k]]['Adda52'].append(round((((dic[key[k]]['Adda52'][i]) * 5) +
                ((dic[key[k]]['Adda52'][i + 1]) * 3) + ((dic[key[k]]['Adda52'][i + 2]) * 2)) / 10,1))

            dic[key[k]]['Dream 11'].append(round((((dic[key[k]]['Dream 11'][i]) * 5) +
                ((dic[key[k]]['Dream 11'][i + 1]) * 3) + ((dic[key[k]]['Dream 11'][i + 2]) * 2)) / 10,1))

            dic[key[k]]['Rummy Circle'].append(round((((dic[key[k]]['Rummy Circle'][i]) * 5) +
                ((dic[key[k]]['Rummy Circle'][i + 1]) * 3) + ((dic[key[k]]['Rummy Circle'][i + 2]) * 2)) / 10,1))



    fields = ['start_date', 'end_date', 'brand', 'audience',
              'FINAL-unaided_brand_awareness_score','FINAL-aided_brand_awareness_score','FINAL-brand_consideration',
              'FINAL-overall_brand_consideration']

    for key in dic:
        for key2 in dic[key]:
            scores = dic[key][key2]

            csvDict = {}
            csvDict['start_date'] = key
            csvDict['end_date'] = key
            csvDict['brand'] = key2
            csvDict['audience'] = game
            csvDict['FINAL-unaided_brand_awareness_score'] = scores[7]
            csvDict['FINAL-aided_brand_awareness_score'] = scores[8]
            csvDict['FINAL-brand_consideration'] = scores[9]
            csvDict['FINAL-overall_brand_consideration'] = scores[10]
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
            for i in range(7,11):
                overallData[day][brand][i-6] = overallData[day][brand][i-6] + scores[i]
        for i in range(1,5):
            overallData[day][brand][i] = round((overallData[day][brand][i] / len(collectByDate[day][brand])), 2)

for day in overallData:
    for brand in overallData[day]:
        csvData = {}
        csvData['start_date'] = day
        csvData['end_date'] = day
        csvData['brand'] = brand
        csvData['audience'] = overallData[day][brand][0]
        csvData['FINAL-unaided_brand_awareness_score'] = overallData[day][brand][1]
        csvData['FINAL-aided_brand_awareness_score'] = overallData[day][brand][2]
        csvData['FINAL-brand_consideration'] = overallData[day][brand][3]
        csvData['FINAL-overall_brand_consideration'] = overallData[day][brand][4]
        output.append(csvData)





filename2 = "Output_Brand_Scores.csv"

with open(filename2, 'w', newline='') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=fields)
    writer.writeheader()
    writer.writerows(output)








