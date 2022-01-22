import win32gui
import win32com.client
import subprocess
import time
import random

shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys('%')

import sys
from PyQt5 import QtCore, QtGui, QtWidgets, uic
 
qtCreatorFile = r"QT\Mainmenu.ui" # Enter file here.
 
Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)
 
chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
url= "yeezysupply.com/product/FV3260"
numProfile = 0 #first one
cmd = chromePath + ' --profile-directory="Profile ' + str(numProfile) + '" ' + url

tasks= 50 #max about 60, use 50 tho
#profile 12 is me
#task=1
#from selenium import webdriver
#options = webdriver.ChromeOptions()
#options.add_argument("user-data-dir=" + r'C:\Users\Moy\AppData\Local\Google\Chrome\User Data')
#options.add_argument("profile-directory=Profile 1")
#driver = webdriver.Chrome(r"F:/Drive/Projects/Session_finder/chromedriver.exe", chrome_options=options)
#driver.get("http://yeezysupply.com")

class MyApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.refreshButton.clicked.connect(self.populateList)
        self.openSession.clicked.connect(self.openListItem)
        self.startButton.clicked.connect(self.launchWindows)
        self.statusbar.showMessage('Message in statusbar.')
        self.websiteBox.setText("yeezysupply.com/product/FV3260")
        
    def launchWindows(self):
        numTasks = int(self.numOfSessions.text()) #max about 60, use 50 tho
        url = self.websiteBox.text()
        print(url)
        for task in range(numTasks):
            print("Opening sesion # %s" % str(task))
            randDelay = random.randint(1,25)
            cmd = cmd = chromePath + ' --profile-directory="Profile ' + str(task) + '" ' + url
            subprocess.Popen(cmd)
            #print(cmd)
            print("Opening new session in: %s" % str(10+randDelay))
            time.sleep(10 + randDelay)
        
    def populateList(self):
        self.listWidget.clear()
        results = []
        top_windows = []
        win32gui.EnumWindows(windowEnumerationHandler, top_windows)
        searchText = self.searchBar.text().lower()
        if searchText == "":
            searchText = "N/A"
        
        for i in top_windows:
            if searchText in i[1].lower():
                results.append(i)
                print(win32gui.GetWindowText(i[0]))
                print(i[0])
                print("Found: " + i[1])
                #win32gui.ShowWindow(i[0], 5)
                #win32gui.SetForegroundWindow(i[0])
            else:
                count = 0
                #print("notepad not found: " + i[1])
        print("Loop complete")
        
        
        if results!= []:
            for result in results:
                print("ID: %s TITLE: %s"%(result[0],result[1]))
        else:
            print("No Results: ")

        for item in results:
            self.listWidget.addItem("✔️ - " +item[1])
            
    def openListItem(self):
        print("OPEN")

        
        
def windowEnumerationHandler(hwnd, top_windows):
    top_windows.append((hwnd, win32gui.GetWindowText(hwnd)))





    
if __name__ == "__main__":    
    app = QtWidgets.QApplication(sys.argv)
    window = MyApp()
    window.setWindowIcon(QtGui.QIcon(r"QT\vodc9otcnpex.png") )
    window.show()
    sys.exit(app.exec_())
    


    time.sleep(20)    
