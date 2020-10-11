from mlmailclassify import DataPrep2
from mlmailclassify import ML_model

root_path = "C:\\Users\\d4an\\OneDrive - Kemira Oyj\\Desktop\\Projects\\ML project\\New_version\\Mails_database\\"


def mail_predict(singl_mail):
    df = DataPrep2.get_clean_data(root_path, singl_mail)
    msg_label = ML_model.classify_mail(df)
    if msg_label is not None:
        return str(msg_label)


# just to teach the model
def train_model():
    DataPrep2.get_clean_data(root_path, save_model=True)
    ML_model.teach_model()
