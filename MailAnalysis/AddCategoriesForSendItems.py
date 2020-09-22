import pandas as pd
from numpy import nan

df = pd.read_excel(r'C:\Users\d4an\OneDrive - Kemira Oyj\Desktop\Projects\EMailAnalysis\WaterExports\WaterExpSent.xlsx')
# df = df.rename(
#     columns={'Unnamed: 0': 'Subject', 'Unnamed: 1': 'Recieving time', 'Unnamed: 2': 'Sender Name', 'Unnamed: 3': 'Body',
#              'Unnamed: 4': 'To', 'Unnamed: 5': 'CC', 'Unnamed: 6': 'Categories'})
a = 'Katarzyna Puzdrowska ; Katarzyna Stachowska; Kornelia Bojanowska ; Luiza; Marlena ; Martyna; Marzena; Nataliia; ' \
    'Oliwia ; Vladyslav ; '

MyCollegues2 = []
MyCollegues = a.split(';')

for Name in MyCollegues:
    MyCollegues2.append(Name.strip())

MyMatch = bool
index = 0
NameList = []
for massage in df['Body']:
    massage = str(massage)
    # splitMassage = massage.split()
    for Name2 in MyCollegues2:
        try:
            if massage.lower().count(Name2.lower()) != 0:
                MyMatch = True
                NameList.append(Name2)
                break
                # print(index + 1, Name2)
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


df = df.to_excel(r'C:\Users\d4an\OneDrive - Kemira Oyj\Desktop\Projects\EMailAnalysis\WaterExports\WaterSentWithCateg.xlsx',
                 index=False)


