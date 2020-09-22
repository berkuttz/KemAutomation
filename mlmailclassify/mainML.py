from mlmailclassify import DataPrep2

root_path = "C:\\Users\\d4an\\OneDrive - Kemira Oyj\\Desktop\\Projects\\ML project\\New_version\\Mails_database\\"

def get_MLpred(singl_mail):

    df = DataPrep2.get_msg(root_path, singl_mail)
    msg_label = DataPrep2.classify_mail(df)
    if msg_label is not None:
        if msg_label[0] is not "Other": print("ML classified it as ", msg_label)


# just to teach the model
def train_model():
    df = DataPrep2.get_msg(root_path)
    DataPrep2.dataset.write_data('cleanlabelcut', df)
    DataPrep2.teach_model()

