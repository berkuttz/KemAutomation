'''
TESTING DECISION TREE
'''

from nltk.corpus import stopwords
from nltk.stem.porter import PorterStemmer
import re
import os
import pandas as pd

from MyModules import datasets

from xgboost import XGBClassifier
from sklearn.naive_bayes import GaussianNB
from sklearn.ensemble import RandomForestClassifier
from sklearn.neighbors import KNeighborsClassifier
from sklearn.tree import DecisionTreeClassifier

dataset = datasets.Dataset_20200521()


def get_clean_data(root_path, singl_mail=False):
    """
    :param root_path: path to nested folder wth mails
    :param singl_mail: when using model passing messege from Outllok as Object
    :return: dataframe
    """

    def textpreproces(Body):
        # remove eerything that are not latters
        # make everything lowercase and keep root of the word
        corpus = []
        review = re.sub('[^a-zA-Z]', ' ', Body)  # delete everything that not letters
        review = review.lower()
        review = review.split()
        ps = PorterStemmer()
        all_stopwords = stopwords.words('english')
        review = [ps.stem(word) for word in review if not word in set(all_stopwords)]
        review = ' '.join(review)
        corpus.append(review)
        return corpus[0]

    def cut_body(body, sender):
        # cut mailbody. Every massage here having whole mailchain in the body.
        bodystr = body.split(sender)[0]
        bodystr.lower()
        with open(r'C:\Users\d4an\PycharmProjects\KemAutomation\mlmailclassify\CutKeyWords.txt') as f:
            for message_cut_part in f:
                bodystr = bodystr.split(message_cut_part.lower())[0]
        return bodystr

    # prepare data and create df in the end.
    PDFbool = []
    body = []
    zfrom = []
    Label = []
    Subject = []

    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # for each folder (that are label) check all files inside. Files are .msg from Outlook
    for folder in os.listdir(root_path):
        messeges = os.listdir(root_path + folder)
        for item in messeges:
            msg = outlook.OpenSharedItem(root_path + folder + "\\" + item)
            # if i am checking one mail (using model and not preparing data for teaching)
            if singl_mail: msg = singl_mail

            # get True if PDF is attached
            pdfbool = False
            for ItemNr in range(1, msg.Attachments.Count + 1):
                if msg.Attachments.Item(ItemNr).FileName[-3:] in ['pdf', 'PDF']:
                    PDFbool.append('PDF')
                    pdfbool = True
                    break
                elif ItemNr == msg.Attachments.Count:
                    PDFbool.append('0')
                    pdfbool = True
            if not pdfbool:
                continue

            bodystr = cut_body(msg.Body, msg.SenderName)
            # save mail attributes
            body.append(textpreproces(bodystr))
            Subject.append(textpreproces(msg.Subject.replace("RE: ", "").replace("FW: ", "")))
            zfrom.append(msg.SenderName)

            if singl_mail:
                data = {'Subject': Subject, 'SenderName': zfrom, 'Body': body,
                        'AttachCount': PDFbool}
                df = pd.DataFrame(data=data)
                return df
            Label.append(folder)

    data = {'Subject': Subject, 'SenderName': zfrom, 'Body': body,
            'AttachCount': PDFbool, 'Label': Label}

    df = pd.DataFrame(data=data)
    return df


def teach_model():
    # test some popular models and store results in df
    def test_model():

        # nested list. Here will be stored results from running models
        results = []
        for features in range(30, 300, 10):
            from sklearn.feature_extraction.text import CountVectorizer
            cv = CountVectorizer(max_features=features)
            X = cv.fit_transform(df.Total).toarray()
            dataset.write_data_pickle('Vectorizer', cv)

            y = df.Label.values

            # Splitting the dataset into the Training set and Test set
            from sklearn.model_selection import train_test_split
            X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.20, random_state=0)

            # loop through models and store results in the list

            for n_estimators in range(10, 150, 5):
                for leaf in range(1,10,1):
                    classifier = RandomForestClassifier(n_estimators=n_estimators, max_depth=leaf)
                    classifier.fit(X_train, y_train)
                    y_pred = classifier.predict(X_test)

                    from sklearn.metrics import accuracy_score
                    results.append([classifier, features, accuracy_score(y_test, y_pred), n_estimators, leaf])

        return results

    # select model with best accurency
    def select_model(model, features):
        print(' Nr of words:', features, 'Accuracy', df_models.iloc[0][2], 'n_estim', df_models.iloc[0][3], 'leaf', df_models.iloc[0][4])
        from sklearn.feature_extraction.text import CountVectorizer
        cv = CountVectorizer(max_features=features)
        X = cv.fit_transform(df.Total).toarray()

        # save CV, load it when classifying single mail
        dataset.write_data_pickle('Vectorizer', cv)
        y = df.Label.values

        # Splitting the dataset into the Training set and Test set
        from sklearn.model_selection import train_test_split
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.20, random_state=0)

        classifier = RandomForestClassifier(n_estimators=df_models.iloc[0][3], max_depth= df_models.iloc[0][4])
        classifier.fit(X_train, y_train)

        # save model
        dataset.write_data_pickle('ML_model', classifier)

    models = [XGBClassifier, GaussianNB, RandomForestClassifier, KNeighborsClassifier, DecisionTreeClassifier]
    # get trained data
    df = dataset.read_data('cleanlabelcut')
    df['Total'] = df[df.columns[:-1]].apply(
        lambda x: ','.join(x.dropna().astype(str)),
        axis=1)

    # best_model - nested list
    best_models = test_model()
    # get DataFrame from nested list
    # looks like this: <class 'xgboost.sklearn.XGBClassifier'>  100  0.981818
    df_models = pd.DataFrame(best_models[0:], columns=['Model', 'CV', 'Accuracy', 'n_estim', 'leaf'])
    df_models = df_models.sort_values('Accuracy', ascending=False)
    # print top 5 models with details
    print(df_models[:5])
    # get best feature ( CV ) and model
    features = int(df_models.iloc[0][1])
    model = df_models.iloc[0][0]

    # return classifier
    select_model(model=model, features=features)


def classify_mail(my_mail, threshold=0.5):
    model = dataset.read_data_pickle('ML_model')
    cv = dataset.read_data_pickle('Vectorizer')

    my_mail['Total'] = my_mail[my_mail.columns[:]].apply(
        lambda x: ','.join(x.dropna().astype(str)),
        axis=1)

    VerfData = cv.transform(my_mail['Total']).toarray()
    # print(model.predict_proba(VerfData))

    if len(VerfData) != 0:
        predicted_proba = model.predict_proba(VerfData)
        print(predicted_proba)
        print(model.classes_)
        for probs in predicted_proba:
            # Iterating over class probabilities
            for i in range(len(probs)):
                if probs[i] >= threshold:
                    # We add the class
                    # print('Probability ', probs[i])
                    return model.classes_[i]
    else:
        return "Trash"
        # return model.predict(VerfData)
