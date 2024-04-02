#  Purpose : Import Weight Scale Data into Garmin Connect Account
#  Author  : Andre Mariano
#  Copyright (c) 2024 Mr Tech Inc. All rights reserved.

# Created : Mar 21st 2024
# Updated : Mar 29th 2024

from garminconnect import Garmin
from datetime import datetime
from timeit import default_timer as timer
import pandas as pd
import json
import sys
from typing import Optional
from pathlib import Path
import re
import email
import imaplib
import os

class FetchEmail():

    connection = None
    error = None

    def __init__(self, mail_server, username, password, folder):
        self.connection = imaplib.IMAP4_SSL(mail_server)
        self.connection.login(username, password)
        self.connection.select(mailbox=folder, readonly=False) # so we can mark mails as read

    def close_connection(self):
        """
        Close the connection to the IMAP server
        """
        self.connection.close()

    def save_attachment(self, msg, download_folder="/tmp"):
        """
        Given a message, save its attachments to the specified
        download folder (default is /tmp)

        return: file path to attachment
        """
        att_path = "No attachment found."
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()

            att_path = os.path.join(download_folder, filename)

            if not os.path.isfile(att_path):
                print("Downloading file " + att_path)
                fp = open(att_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
        return att_path

    def fetch_unread_messages(self):
        """
        Retrieve unread messages
        """
        emails = []
        (result, messages) = self.connection.search(None, 'UnSeen')
        if result == "OK":
            for message in messages[0].split():
                try: 
                    ret, data = self.connection.fetch(message,'(RFC822)')
                except:
                    #print("No new emails to read.")
                    self.close_connection()
                    exit()

                msg = email.message_from_bytes(data[0][1])
                if isinstance(msg, str) == False:
                    emails.append(msg)
                response, data = self.connection.store(message, '+FLAGS','\\Seen')

            return emails

        self.error = "Failed to retreive emails."
        return emails

    def parse_email_address(self, email_address):
        """
        Helper function to parse out the email address from the message

        return: tuple (name, address). Eg. ('John Doe', 'jdoe@example.com')
        """
        return email.utils.parseaddr(email_address)

class Timer:

    def __init__(self):        
        self.startTimer()

    def startTimer(self):
        self.start = timer()

    def endTimer(self):
        self.end = timer()
        print('Elapsed Time: %.2f seconds' %(self.end - self.start))
        print('')
        self.startTimer()    

class ImportBodyComposition:

    def __init__(self):
        self.loadPropertiesFile()
        self.loadColumnsMapping()
        self.initBodyCompositionDF()

        self.lastDate = self.getPropertyValue('data', 'lastDate', None)

    def loadPropertiesFile(self):
        if (len(sys.argv) > 1) and (sys.argv[1]):
            self.propertiesFile = sys.argv[1] 
        else:   
            self.propertiesFile = '//app/app//garminconnect.json'  

        with open(self.propertiesFile, encoding="utf8") as jsonFile:
            jsonContents = jsonFile.read()

        self.jsonProperties = json.loads(jsonContents) 

    def downloadAttachments(self):

        mailHost = self.getPropertyValue('imap', 'host', None)
        mailUserName = self.getPropertyValue('imap', 'userName', None)
        mailPassword = self.getPropertyValue('imap', 'password', None)
        mailFolder = self.getPropertyValue('imap', 'folder', None)
        weightFilesFolder = self.getPropertyValue('data', 'weightFilesFolder', None)

        f = FetchEmail(mailHost, mailUserName, mailPassword, mailFolder)

        emails = f.fetch_unread_messages()

        for msg in emails:
            f.save_attachment(msg, weightFilesFolder)                  

    def loadColumnsMapping(self):
        # Load Columns Mapping Info
        self.mapInfo = {}
        for key in self.jsonProperties['data'] ['Mapping']:
            self.mapInfo[key] = self.jsonProperties['data'] ['Mapping'][key]

    def countBodyComposition(self):
        return len(bc.BodyComposition)
    
    def initBodyCompositionDF(self):
        self.BodyComposition = pd.DataFrame([], columns=['timestamp', 'weight', 'percent_fat', 'percent_hydration', 'visceral_fat_mass', 'bone_mass', 'muscle_mass', 'basal_met', 'active_met', 'physique_rating', 'metabolic_age', 'visceral_fat_rating', 'bmi'])

    def getPropertyValue(self, section, property, default: Optional[any] = None):
        try:
            if   (property in self.jsonProperties[section]):
                propertyValue = self.jsonProperties[section] [property]
            else:
                propertyValue = default 
        except Exception as e:
            propertyValue = default

        return propertyValue
    
    def updatePropertiesFile(self):
        with open(self.propertiesFile, "w") as jsonFile:
            json.dump(self.jsonProperties, jsonFile, indent=4)

    def connectGarmin(self):
        email = self.getPropertyValue('garmin', 'email')
        password = self.getPropertyValue('garmin', 'password')

        if  not email or not password:
            raise SystemExit("Garmin Credentials was not informed.")
        
        print("Log on Garmin Account: " + email)    
        self.client = Garmin(email, password)
        self.client.login()  
        print("  Garmin User Full Name: " + self.client.full_name)

    def validMandatoryValues(self, row):
        try:
            boolTest = True

            for key in row:
                if  row[key]: 
                    hasValue = True
                else:
                    hasValue = False

                try:
                    isMandatory = (key in self.mapInfo) and (self.mapInfo[key]['mandatory'] == 'True')
                except Exception as e:
                    isMandatory = False

                boolTest = boolTest and (hasValue or not isMandatory)

                if  not boolTest:
                    break
                
            return boolTest
        except ValueError as e:
            print("Conversion Error: ", e)
            raise
        except Exception as e:
            print("Error retrieving row definition: ", e)
            raise               

    def getMappingColumnValue(self, row, columnName):
        try:
            value = None
            if (columnName in self.mapInfo) and ('name' in self.mapInfo[columnName]) and (self.mapInfo[columnName]['name'] in row):
                rowValue = row[self.mapInfo[columnName]['name']]
                if  ('type' in self.mapInfo[columnName]):
                    colType = self.mapInfo[columnName]['type']
                else:
                    colType = "value"

                regex = None    

                if colType == "weight":
                    regex = r'^([\d.]+)'
                elif colType == "kcal":
                    regex = r'(\d+\.?\d*)'
                elif colType == "percent":
                    regex = r'(\d+\.\d+)'
                elif colType == "value":
                    value = float(rowValue)
                else:
                    value = rowValue

                if regex:
                    valueStr = str(rowValue)
                    match = re.search(regex, valueStr)
                    if (match):
                        value = float(match.group(1))
            else:
                value = None        
        except ValueError as e:
            value = None
        except Exception as e:
            value = None 

        return value

    def getMappingRowValues(self, row):
        rowValues = {}
        
        for key in self.BodyComposition.columns:
            rowValues[key] = self.getMappingColumnValue(row, key)

        return rowValues

    def processWeightFile(self, weightFile, fileName):
        filterByUser = self.getPropertyValue('data', 'filterByUser', 'False') == 'True'
        user = self.getPropertyValue('data', 'user')
        dateTimeFormat = self.getPropertyValue('data', 'dateTimeFormat')
        sortData = self.getPropertyValue('data', 'sortData', 'False') == 'True'
        
        if  not dateTimeFormat:
            raise SystemExit("Date Time Format was not informed.")

        if  filterByUser and not user:
            raise SystemExit("Filter user was not informed.")  

        print('  Weight File: ' + fileName)    

        # Load CSV Weight File
        InputData = pd.read_csv(weightFile)
        
        # Convert timestamp to Datetime
        InputData[self.mapInfo['timestamp']['name']] = pd.to_datetime(InputData[self.mapInfo['timestamp']['name']], format = dateTimeFormat)

        if  (sortData):
            if  (filterByUser):
                InputData = InputData.sort_values(by = [self.mapInfo['userName']['name'], self.mapInfo['timestamp']['name']], ascending = [True, True])
            else:
                InputData = InputData.sort_values(by = [self.mapInfo['timestamp']['name']], ascending=[True])

        print('    %4d rows found.' %len(InputData))
        insertedItems = 0 
        for index, rowInput in InputData.iterrows():            
            if  ((not filterByUser) or self.getMappingColumnValue(rowInput, 'userName') == user):
                rowOutput = self.getMappingRowValues(rowInput)
                dateISOformat = rowOutput['timestamp'].date().isoformat()
                if  (not self.lastDate or (dateISOformat >= self.lastDate)):                
                    findRecord = self.BodyComposition.loc[self.BodyComposition['timestamp'] == rowOutput['timestamp']]
                    if  (self.validMandatoryValues(rowOutput) and findRecord.empty):
                        self.BodyComposition.loc[len(self.BodyComposition)] = rowOutput 
                        insertedItems += 1

        # Sort dataframe by timestamp
        self.BodyComposition = self.BodyComposition.sort_values(by = ['timestamp'], ascending=[True])
        
        discardedItems = len(InputData) - insertedItems

        print('    %4d rows discarded.' %(discardedItems))
        print('    %4d rows loaded.' %(insertedItems))    
        print('')

    def loadDataFromWeightFilesFolder(self):
        weightFilesFolder = self.getPropertyValue('data', 'weightFilesFolder', None)
        fileMask = self.getPropertyValue('data', 'fileMask', '*.csv')

        if  not weightFilesFolder:
            raise SystemExit("Weight File folder was not informed.")        

        print('Processing ' + weightFilesFolder + '\\' + fileMask + ' files...')

        if  (self.lastDate):
            print('  Last Imported Date: ' + self.lastDate)
        print('')
        files = Path(weightFilesFolder).glob(fileMask)
        for file in files:
            self.processWeightFile(file._str_normcase, file.name)        

    def loadDataOnGarminConnect(self):
        callAPI = self.getPropertyValue('data', 'callAPI', 'False') == 'True'
        deleteOldData = self.getPropertyValue('data', 'deleteOldData', 'False') == 'True'
        
        if  (callAPI):
            print("Loading Body Composition Data into Garmin Connect...")        
            print('')

            dates = []
            loadedItems = 0
            for index, row in self.BodyComposition.iterrows():
                dateISOformat = row['timestamp'].date().isoformat()
                
                if  not (dateISOformat in dates):
                    print('  ' + dateISOformat)
                    dates.append(dateISOformat)

                    self.jsonProperties['data'] ['lastDate'] = dateISOformat 

                    daily_weigh_ins = self.client.get_daily_weigh_ins(dateISOformat)
                    if  (deleteOldData and daily_weigh_ins['dateWeightList']):
                        print("    Deleting weight ins") 
                        for w in daily_weigh_ins['dateWeightList']:
                            self.client.delete_weigh_in(w["samplePk"], dateISOformat)

                self.client.add_body_composition(timestamp = row['timestamp'].isoformat(), 
                                                 weight = row['weight'],
                                                 percent_fat = row['percent_fat'],
                                                 percent_hydration = row['percent_hydration'],
                                                 visceral_fat_mass = row['visceral_fat_mass'],
                                                 bone_mass = row['bone_mass'],
                                                 muscle_mass = row['muscle_mass'],
                                                 basal_met = row['basal_met'],
                                                 active_met = row['active_met'],        
                                                 physique_rating = row['physique_rating'],  
                                                 metabolic_age = row['metabolic_age'],  
                                                 visceral_fat_rating = row['visceral_fat_rating'],
                                                 bmi = row['bmi'])
                
                print("    Body Composition inserted: " + row['timestamp'].strftime("%H:%M:%S")) 

                loadedItems += 1  

            print('')
            print('    %4d weight ins was imported on Garmin Account.' %loadedItems)
            print('')

if __name__ == "__main__":
    print("Import Body Composition Data into Garmin Connect Account")
    print("Copyright (c) 2024 Mr Tech Inc. All rights reserved.")
    print("")

    bc = ImportBodyComposition() 
    t = Timer()  

    bc.loadPropertiesFile()

    t.startTimer()
    bc.downloadAttachments()
    t.endTimer()    

    bc.loadDataFromWeightFilesFolder()
    t.endTimer()

    if  (bc.countBodyComposition() > 0):
        bc.connectGarmin()
        t.endTimer()
        
        bc.loadDataOnGarminConnect()
        t.endTimer()

        bc.updatePropertiesFile()
    else:
        print('No loaded data to import on Garmin Account')