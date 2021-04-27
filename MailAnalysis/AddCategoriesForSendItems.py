import pandas as pd

df = pd.read_excel(r'C:\Users\d4an\OneDrive - Kemira Oyj\Desktop\TM KU Sent items.xlsm')
ExcelWithTeam = pd.read_excel(r'C:\Users\d4an\OneDrive - Kemira Oyj\Desktop\TankMonitorTeam.xlsx')

ExcelWithTeam = ExcelWithTeam.dropna()

MyCollegues = []


for Name in ExcelWithTeam.Person:
    MyCollegues.append(Name.strip())

MyMatch = bool
index = 0
NameList = []
for massage in df['Body']:
    massage = str(massage)
    # splitMassage = massage.split()
    for Name2 in MyCollegues:
        try:
            if massage.lower().count(Name2.lower()) != 0:
                MyMatch = True
                NameList.append(Name2)
                break

        except:
            break
    if not MyMatch:
        NameList.append("Unknown")
        # print(massage)
        # print('for this one nothing was found nr ', index + 1)
    MyMatch = False
    index += 1

df['Categor'] = NameList

# data = {'Index': index, 'WhoSentThis': NameList}
# lengthOfMasg = pd.DataFrame(data=data, index=range(0, len(df)))


df = df.to_excel(r'C:\Users\d4an\OneDrive - Kemira Oyj\Desktop\TankMonitorSent.xlsx',
                 index=False)


