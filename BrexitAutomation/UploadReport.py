from MyModules import SAP_Class
from MyModules import WebAutomation
from MyModules import utils


'''
requirements
OS - English
printer - Foxit
'''

if __name__ == '__main__':

    '''
    Define variables 
    '''
    UserLogin = utils.global_variable().login()
    variant = "BRXUKEUVAT"
    layout = "brexit"
    file_name = "export12.XLSX"

    # run the report in SAP
    SAP = SAP_Class.VladSAP()
    SAP.download_report_ZL06O(variant, layout, file_name)
    SAP.close_window()

    # upload report to Web Site
    WEB = WebAutomation.JenkarPortal()
    WEB.logingtoSIte()
    WEB.upload_excel(file_name)
    raise SystemExit(0)

