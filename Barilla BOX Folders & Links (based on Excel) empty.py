import sys
import boxsdk
from boxsdk import Client, OAuth2, JWTAuth

import boxsdk.object.comment
import boxsdk.object.collection
import boxsdk.object.collaboration
import boxsdk.object.collaboration_whitelist_entry
import boxsdk.object.collaboration_whitelist_exempt_target
import boxsdk.object.collaboration_allowlist_entry
import boxsdk.object.collaboration_allowlist_exempt_target
import boxsdk.object.device_pinner
import boxsdk.object.enterprise
import boxsdk.object.events
import boxsdk.object.event
import boxsdk.object.file
import boxsdk.object.file_version
import boxsdk.object.file_version_retention
import boxsdk.object.folder
import boxsdk.object.group
import boxsdk.object.group_membership
import boxsdk.object.legal_hold
import boxsdk.object.legal_hold_policy
import boxsdk.object.legal_hold_policy_assignment
import boxsdk.object.metadata
import boxsdk.object.metadata_cascade_policy
import boxsdk.object.metadata_template
import boxsdk.object.recent_item
import boxsdk.object.retention_policy
import boxsdk.object.retention_policy_assignment
import boxsdk.object.search
import boxsdk.object.storage_policy
import boxsdk.object.storage_policy_assignment
import boxsdk.object.task
import boxsdk.object.task_assignment
import boxsdk.object.terms_of_service
import boxsdk.object.terms_of_service_user_status
import boxsdk.object.user
import boxsdk.object.upload_session
import boxsdk.object.watermark
import boxsdk.object.web_link
import boxsdk.object.webhook

from math import ceil
import json
from PySide2 import QtWidgets, QtGui, QtCore
from PySide2.QtCore import (QCoreApplication, QMetaObject, QRect, Qt)
from PySide2.QtGui import (QFont)
from PySide2.QtWidgets import *
import pandas as pd


###############################################
# CREATE APP
###############################################
class Ui_Dialog(object):
    def setupUi(self, Dialog):
        if Dialog.objectName():
            Dialog.setObjectName(u"Dialog")
        Dialog.resize(403, 370)
        Dialog.setStyleSheet(u"background: #3a3a47\n"
                             "")
        font = QFont()
        font.setBold(True)
        font.setUnderline(False)
        font.setWeight(75)
        font.setStrikeOut(False)


        self.pushButton2 = QPushButton(Dialog)
        self.pushButton2.setObjectName(u"pushButton")
        self.pushButton2.setGeometry(QRect(50, 70, 300, 50))
        font1 = QFont()
        font1.setPointSize(10)
        font1.setBold(True)
        font1.setUnderline(False)
        font1.setWeight(75)
        font1.setStrikeOut(False)
        self.pushButton2.setFont(font1)
        self.pushButton2.setStyleSheet(u"QPushButton {\n"
                                       "	color: #FFFFFF !important;\n"
                                       "   font: 12pt 'Century Gothic';\n"
                                       "	text-transform: uppercase;\n"
                                       "	text-decoration: none;\n"
                                       "	background: #1d1c21;\n"
                                       "	padding: 0px;\n"
                                       "	border: 4px solid #494949 !important;\n"
                                       "	display: inline-block;\n"
                                       "	transition: all 0.4s ease 0s;\n"
                                       "}\n"
                                       "QPushButton:hover {\n"
                                       "	color: #ffffff !important;\n"
                                       "	background: #f6b93b;\n"
                                       "	border-color: #f6b93b !important;\n"
                                       "	transition: all 0.4s ease 0s;\n"
                                       "}\n"
                                       "")

        self.lineEdit = QLineEdit(Dialog)
        self.lineEdit.setObjectName(u"lineEdit")
        self.lineEdit.setGeometry(QRect(50, 30, 300, 30))
        self.lineEdit.setFont(font)
        self.lineEdit.setStyleSheet(u"color: #494949 !important;\n"
                                    "text-transform: uppercase;\n"
                                    "text-decoration: none;\n"
                                    "font: 12pt 'Century Gothic';\n"
                                    "color: #FFFFFF;\n"
                                    "background: #1d1c21;\n"
                                    "border: none;\n"
                                    "display: inline-block;")
        self.lineEdit.setAlignment(Qt.AlignCenter)

        self.pushButton = QPushButton(Dialog)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(50, 280, 300, 61))
        self.pushButton.setFont(font1)
        self.pushButton.setStyleSheet(u"QPushButton {\n"
                                      "	color: #FFFFFF !important;\n"
                                      "   font: 12pt 'Century Gothic';\n"
                                      "	text-transform: uppercase;\n"
                                      "	text-decoration: none;\n"
                                      "	background: #A9742A;\n"
                                      "	padding: 0px;\n"
                                      "	border: 4px solid #494949 !important;\n"
                                      "	display: inline-block;\n"
                                      "	transition: all 0.4s ease 0s;\n"
                                      "}\n"
                                      "QPushButton:hover {\n"
                                      "	color: #ffffff !important;\n"
                                      "	background: #f6b93b;\n"
                                      "	border-color: #f6b93b !important;\n"
                                      "	transition: all 0.4s ease 0s;\n"
                                      "}\n"
                                      "")
        self.progressBar = QProgressBar(Dialog)
        self.progressBar.setObjectName(u"progressBar")
        self.progressBar.setGeometry(QRect(50, 230, 300, 40))
        self.progressBar.setFont(font)
        self.progressBar.setStyleSheet(u"QProgressBar {\n"
                                       "	color: #FFFFFF !important;\n"
                                       "   font: 12pt 'Century Gothic';\n"
                                       "	text-transform: uppercase;\n"
                                       "	text-decoration: none;\n"
                                       "	background: #1d1c21;\n"
                                       "	border: 4px solid #494949 !important;\n"
                                       "	display: inline-block;\n"
                                       "	text-align: center;\n"
                                       "}\n"
                                       "\n"
                                       "QProgressBar::chunk {\n"
                                       "   background-color: #A9742A;\n"
                                       "   width: 20px;\n"
                                       "	text-align: center;\n"
                                       "}")
        self.progressBar.setValue(0)
        self.retranslateUi(Dialog)

        QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(QCoreApplication.translate("Dialog", u"BOX folders creation", None))
        self.lineEdit.setText("")
        self.lineEdit.setPlaceholderText(QCoreApplication.translate("Dialog", u"folder id on box", None))
        self.pushButton.setText(QCoreApplication.translate("Dialog", u"make awesome", None))
        self.pushButton2.setText(QCoreApplication.translate("Dialog", u"choose file", None))


###############################################
# OPEN APP
###############################################
app = QtWidgets.QApplication(sys.argv)

# init
Dialog = QtWidgets.QDialog()
ui = Ui_Dialog()
ui.setupUi(Dialog)
Dialog.show()


class jiraUpdate:
    ###############################################
    # IMPORT TAB DATA
    ###############################################
    def chooseFile():
        global new_folders
        global firstCol
        global path
        filename = QFileDialog.getOpenFileName()
        path = filename[0]
        new_folders = pd.read_excel(path)
        new_folders.drop_duplicates(inplace=True)
        listCol = new_folders.columns.tolist()
        firstCol = listCol[0]

    def jiraData():
        mainID = ui.lineEdit.text()
        statusBarProgress = ui.progressBar.setValue(0)

        ###############################################################################
        # AUTHORIZATION
        ###############################################################################

        auth = JWTAuth(
            enterprise_id="",
            client_id="",
            client_secret="",
            rsa_private_key_data=bytes(
                "", 'ascii'),
            rsa_private_key_passphrase=bytes("", 'ascii'),
            jwt_key_id="",
        )

        client2 = Client(auth)
        service_account = client2.user().get()
        access_token = auth.authenticate_instance()

        app_user = client2.user(user_id='')

        auth2 = JWTAuth(
            enterprise_id=None,
            user=app_user,
            client_id="",
            client_secret="",
            rsa_private_key_data=bytes(
                "", 'ascii'),
            rsa_private_key_passphrase=bytes("", 'ascii'),
            jwt_key_id="",
        )
        access_token2 = auth2.authenticate_user()
        client = Client(auth2)

        ##############################################################################
        # FOLDERS CREATION / LINKS CREATION / DEPLOY
        ##############################################################################
        statusBar = 0
        i = 0
        statusBarElement = 100 / len(new_folders)

        for new_folder in new_folders[firstCol]:
            QtCore.QCoreApplication.processEvents()

            subfolder = client.folder(folder_id=mainID).create_subfolder(str(new_folder))
            urlsubfolder = client.folder(folder_id=subfolder.id).get_shared_link(access='Open')
            new_folders.loc[i,"MAIN FOLDER"] = urlsubfolder

            i += 1
            statusBar = statusBar + statusBarElement
            statusBarProgress = ui.progressBar.setValue(statusBar)


        new_folders.to_excel(path,index=True)

start = ui.pushButton.clicked.connect(jiraUpdate.jiraData)
start2 = ui.pushButton2.clicked.connect(jiraUpdate.chooseFile)

# Main loop
sys.exit(app.exec_())
