
import re
from MyModules import utils, datasets
import os

class OulAutom:

    def add_attachment(self, messeges, filename):
        from MyModules import SAP_Class
        path = utils.global_variable().file_path()
        for att in messeges.Attachments:
            if att.FileName[-3:] in ['pdf', 'PDF']:
                # downdload file
                if att.FileName in os.listdir(path): os.remove(path + att.FileName)
                att.SaveAsFile(path + att.FileName)
                # rename file
                if filename in os.listdir(path): os.remove(path + filename)
                os.rename(path + att.FileName, path + filename)
                # get all deliveries
                deliveries = re.findall(r'\d+', messeges.Subject)
                # for each deivery go to Invoice and attach the file
                SAP = SAP_Class.VladSAP()
                for deliv in deliveries:
                    if deliv.startswith(('202', '83', '85')) and len(deliv) > 7:
                        SAP.open_del_03(deliv)
                        SAP.del_to_inv(change=False, output=False)
                        SAP.add_attachment(path, filename)
                SAP.close_window()
                os.remove(str(path) + '\\' + str(filename))
                break

    def add_cost_BradUS(self,messeges):
        for att in messeges.Attachments:
            if att.FileName[-3:] in ['pdf', 'PDF']:
                pass

