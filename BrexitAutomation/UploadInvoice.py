from MyModules import SAP_Class
from MyModules import WebAutomation


if __name__ == '__main__':
    SAP = SAP_Class.VladSAP()
    SAP.get_fake_del()
    # # todo get some list of shipments/delvieries
    SAP.open_del_03("83616510")
    SAP.del_to_inv(filename="83616510 Inv")
    SAP.close_window()

    WEB = WebAutomation.JenkarPortal()
    WEB.logingtoSIte()
    WEB.upload_invoice("O19574920")
    print("the end")
    raise SystemExit(0)


