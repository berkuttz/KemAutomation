import os


UserLogin = os.getlogin()
if UserLogin == "d4an":
       path = 'C:\\Users\\d4an\\OneDrive - Kemira Oyj\\Desktop\\Projects\\ML project\\'
else:
       path = 'C:\\Users\\' + UserLogin + '\\Downloads\\'
