from win32com.client import Dispatch
import re
from collections import Counter
import time
from datetime import datetime
from MyModules import utils, datasets
from OutlookAutomatization import MainAutom
Oul = MainAutom.OulAutom()

## imports for ML to EXE
# import sklearn.utils._cython_blas
# import sklearn.neighbors.typedefs
# import sklearn.neighbors.quad_tree
# import sklearn.tree
# import sklearn.tree._utils

class MailCategorize:

    def __init__(self):

        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

        folder = outlook.Folders.Item("Shipments Export")

        self.inbox = folder.Folders.Item("Inbox")
        self.outbox = folder.Folders.Item("Sent Items")
        self.oldmails = self.loadoldmails()
        self.sentitems = self.loadsentitems()
        dataset = datasets.Dataset_20200521()
        df = dataset.read_data('ShipExpTeam')
        self.team = df
        # self.team = self.updateTeam(df)
        self.exludedMails = ['Water Export Orders', 'Orders EMEA CIS', 'EMEA SAP SD Key Users', 'Orders Acrylamide',
                             'Orders Russia Gdansk', 'Orders Oil & Gas', 'Orders Export', 'GTC EMEA']

    # update people who is working today. If nor add their backup
    def updateTeam(self, df):

        for index in range(len(df)):
            if df.StartOff[index] <= datetime.now() <= df.EndOff[index]:
                df.at[index, 'Category'] = df.Backup[index]
        return df

    def getrecentperson(self, mailbody):
        parsedmail = mailbody.split("shipments export")
        for i in range(len(parsedmail)):

            for index in range(len(self.team)):
                try:
                    if 'the exporter of the products covered by ' in parsedmail[i]: continue
                    if self.team['Text in mail'][index].lower() in parsedmail[i]:
                        return self.team.Category[index]
                except:
                    continue

    # load old messages into memory
    def loadoldmails(self):
        oldmails = []
        try:
            inbmsg = self.inbox.Folders.Item("CLOSED").Items
            inbmsg.Sort("[ReceivedTime]", True)
            for messeges in inbmsg:
                try:
                    oldmails.append(messeges)
                    if len(oldmails) > 9000: break
                except:
                    continue
            print('old data were loaded ')
            return oldmails
        except:
            print("RESTART OUTLOOK")
            input(" ")
            raise SystemExit(0)

    def loadsentitems(self):
        sentitems = []
        try:
            senitems = self.outbox.Items
            senitems.Sort("[ReceivedTime]", True)
            for messeges in senitems:
                try:
                    sentitems.append(messeges)
                    if len(sentitems) > 500: break
                except:
                    continue
            print('sent items were loaded')
            return sentitems
        except:
            print("RESTART OUTLOOK")
            SystemExit(0)

    def getcategor(self, catlist):
        # get most frequent categor
        if catlist:
            print("Previous meeseges were for: ", catlist)
            catlist = filter(None, catlist)
            c = Counter(catlist)
            return c.most_common(1)[0][0]

    def getrefnrs(self, reflist):
        # get all references from
        templist = []
        for nr in reflist:
            if not str(nr).startswith('4') and nr not in templist:
                templist.append(nr)
        return str(templist)

    def rhenusmrn(self):
        # MRN from Rhenus
        for messeges in self.inbox.Items:
            listo_of_users = ['Vlad', 'Luiza', 'Martyna']
            if messeges.UnRead and messeges.SenderName == 'documents@cesped.it' and \
                    messeges.Categories == '' and not messeges.Subject.startswith("_mrn"):
                rhenusref = re.findall(r'\d+', messeges.Subject)
                foundnr = []
                for oldmail in self.oldmails:
                    if rhenusref[0] in oldmail.Subject:
                        for word in oldmail.Subject.replace("+", " ").replace(",", " ").split():
                            if str(word).startswith(('202', '83', '85')) \
                                    and len(word) > 7: foundnr.append(word)
                refnr = self.getrefnrs(foundnr)
                if refnr: messeges.Subject = '_mrn ' + str(messeges.Subject) + ' ' + refnr.replace("'", "").replace(
                    "[", "").replace("]", "").replace("\\", "")
                messeges.Save()
                time.sleep(2)

            elif messeges.Subject.startswith("_mrn") and messeges.UnRead and \
                    messeges.Categories in listo_of_users:
                cleanedSubject = re.sub('[^0-9]', ' ', messeges.Subject)
                cleanedSubject = cleanedSubject.strip()
                filename = "_mrn " + cleanedSubject + ".pdf"
                Oul.add_attachment(messeges, filename)
                messeges.UnRead = False
                messeges.Save()

    # add categories
    def categorize(self):
        for messeges in self.inbox.Items:
            foundcat = []
            refnumbers = ''
            MailSubject = messeges.Subject.replace("+", " ").replace \
                ("/", " ").replace("_", " ").replace(",", " ").split()
            try:
                if messeges.UnRead and messeges.Categories == '' and messeges.SenderName not in self.exludedMails:
                    # self.get_MLpred(messeges)
                    categ = self.getrecentperson(messeges.Body.lower())
                    if not categ:
                        if "kemira shipping advice n" in str(messeges.Subject).lower():
                            categ = "Oliwia"
                        else:
                            # if there is no recent colegue in mailchain
                            for word in MailSubject:
                                if str(word).startswith(('20', '83', '85', '45', '100')) and 11 > len(
                                        word) > 7: refnumbers = word
                        if not refnumbers:
                            for word in str(messeges.Body[:20]).split():
                                if word.startswith(('20', '83', '85')) and 11 > len(word) > 7: refnumbers = word
                        if refnumbers:
                            if not categ:
                                for oldmail in self.oldmails:
                                    if refnumbers in oldmail.Subject:
                                        foundcat.append(oldmail.Categories)
                                categ = self.getcategor(foundcat)
                            if not categ:
                                for oldmail in self.sentitems:
                                    if refnumbers in oldmail.Subject:
                                        categ = self.getrecentperson(oldmail.Body.lower())
                    # if some category was found
                    if categ:
                        print("Subject: ", messeges.Subject)
                        messeges.Categories = categ
                        messeges.Save()
                        time.sleep(2)
                        print('***Category was assign to ', messeges.Categories, "\n")
            except:
                import sys
                # print("Error occurred ", sys.exc_info())
                continue

    def ML_labels(self):
        from mlmailclassify import mainML
        for messeges in self.inbox.Items:
            if messeges.UnRead and messeges.Categories == "Vlad" and \
                    not str(messeges.Subject).startswith("_bl") and messeges.SenderName not in self.exludedMails:

                label = mainML.mail_predict(messeges)
                print(messeges.Subject)
                print("***ML think this is", label, '\n')

                if label == "Final_BL":
                    temp_result = re.search("20[\d]{7}|83[\d]{6}|(85[\d]{6})", messeges.Subject)
                    del_nr = temp_result.group(0)
                    filename = "_bl " + del_nr + ".pdf"
                    Oul.add_attachment(messeges, filename)
                    messeges.Subject = "_bl " + str(messeges.Subject)
                    messeges.Save()

                if label == "Draft_BL" and not messeges.IsMarkedAsTask:
                    messeges.MarkAsTask(4)
                    messeges.Save()


if __name__ == '__main__':
    mailclasss = MailCategorize()
    for j in range(9):
        if j > 0:
            # update old mails
            mailclasss.loadoldmails()
            mailclasss.loadsentitems()
        for i in range(30):
            mailclasss.rhenusmrn()
            mailclasss.ML_labels()
            mailclasss.categorize()
            print("Coffee break")
            time.sleep(120)
