import datetime as dt
from datetime import timedelta
from datetime import date
import datetime
import os
import time
from tkinter import *
from typing import final
import requests  # gives access to get/post methods
import win32com.client as client  # gives access to outlook
import openpyxl
import csv
from tabulate import tabulate
import sys



import pandas as pd
import win32timezone
from pandas import ExcelWriter
import pprint
import pyodbc
import openpyxl


def SAPlanningmain(today,currentdirectory):
    today1 = datetime.datetime.now().strftime("%d-%m-%Y")
    print("\nInitializing SA Planning Alert... \n")
    def dataLog(data):
        os.chdir(r'\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\PM Shift\DHP1\SchedulingCompliance\\' + today + ' Compliance Uploads logs')
        name = today + " SA Planning logs.txt"
        time_now = datetime.datetime.now()
        str_to_write = str(time_now) + ": " + data
        # open file
        log_file = open(name, "a")
        # write into file
        log_file.write(str_to_write + "\n\n")
    
    pd.set_option('display.max_rows', 500)
    pd.set_option('display.max_columns', 500)
    pd.set_option('display.width', 1000)

    weekday = datetime.datetime.today().weekday()


    are_we_testing = False


    dataSet = {
    }

    table_template = [
        "/md |Country|Scheduler|First_Forecast|First_SA_Update|First_SA_Report|First_SLA|Second_Forecast|Second_SA_Update|Second_SA_Report|Second_SLA|\n",
        "|-|\n"
    ]

    wave_report = {
        "first": 0,
        "second": 0
    }


    path_date = datetime.datetime.now().strftime("%Y-%m-%d")
    task_date = datetime.datetime.now().strftime("%d-%m-%Y")
    path_to_task_list = fr"\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\PM Shift\DHP1\TaskList\TasklistGeneration\{path_date}" + "\shift_config.xlsx"




    #########GLOBAL VARIABLES########################
    startTime = datetime.datetime.now().time()


    URL = ""
    URL_management = ""


    dataLog("Preparing various date strings (dd/mm/yyy, dd-mm-yyy, yyyy-mm-dd")
    today_second_format = datetime.datetime.now().strftime("%Y-%m-%d")
    slashToday = datetime.datetime.now().strftime("%d/%m/%Y")
    backwordsToday = (datetime.datetime.now() + dt.timedelta(days=1)).strftime("%Y-%m-%d")
    emailDate = datetime.datetime.now().strftime("%d-%m-%Y")


    def testing(isOn):
        url_list = []
        if isOn:
            URL = "https://hooks.chime.aws/incomingwebhooks/c9b2223d-ea73-4785-af64-e722e204ddd4?token=NTdGV3NOZFl8MXxULV85ZjN3R0FRTzRjN0hYV2NTVXJLbFBpZUhvQjVJX1dhanJwRnRfQnhn"
            URL_management = "https://hooks.chime.aws/incomingwebhooks/c9b2223d-ea73-4785-af64-e722e204ddd4?token=NTdGV3NOZFl8MXxULV85ZjN3R0FRTzRjN0hYV2NTVXJLbFBpZUhvQjVJX1dhanJwRnRfQnhn"
            url_list.append(URL)
            url_list.append(URL_management)
            return url_list
        else:
            URL = "https://hooks.chime.aws/incomingwebhooks/6f6fd4d1-4410-4503-810f-a3e320a25f46?token=Wk9PYmxHMnB8MXxlNFQ5NTFSdXpLRmhxZGNzS3F2VlJWTzFvZV90MF9acWJKa2IyVXRyam9V"
            URL_management = "https://hooks.chime.aws/incomingwebhooks/a8be2f2f-33f6-4f98-97b2-3c30d3ac288a?token=b0VISFBwV3J8MXxkb2RGbDRaNlpmSERTME45ME1mTnFnOVlPZlZQcEpLWVc0V0NWR25OWFhz"
            url_list.append(URL)
            url_list.append(URL_management)
            return url_list


    firstWave = 13
    secondWave = 14
    dataLog("First wave:  - " + str(firstWave))
    dataLog("Second wave:  - " + str(secondWave))
    firstWaveTime = datetime.datetime.now().replace(hour=14, minute=2, second=0)
    firstWaveTimeEnd = datetime.datetime.now().replace(hour=14, minute=6, second=0)
    secondWaveTime = datetime.datetime.now().replace(hour=19, minute=2, second=0)
    secondWaveTimeEnd = datetime.datetime.now().replace(hour=19, minute=6, second=0)



    # GET DATA FROM EXCEL FILES

    dataLog("Getting data from config.xlsx + preparing variables.")
    os.chdir(currentdirectory)
    configFile = pd.ExcelFile("data_config.xlsx", engine='openpyxl')
    cfg = configFile.parse(0)
    forecasting_folder = cfg.iloc[3, 1]
    sa_update_folder = cfg.iloc[4, 1]
    sa_report_folder = cfg.iloc[5, 1]
    search_seconds = int(cfg.iloc[8, 1])
    username = str(os.getenv('username')) + "@amazon"+str(cfg.iloc[9, 1])
    listOfCountries = ["DE","ES","IT","FR"]
    
    if weekday == 5:
        print("\nIt's Saturday - removing DE from the list of countries. . .")
        listOfCountries.remove("DE")

    for country in listOfCountries:

        if country == "nan":
            listOfCountries.remove(country)

    for country in listOfCountries:
        dataSet[country] = {
            "user": "asd",
            "forecast": {
                "received1": 0,
                "received2": 0
            },
            "saUpdate": {
                "received1": 0,
                "received2": 0
            },
            "saReport": {
                "received1": 0,
                "received2": 0,
            },
            "chimePing": {
                "first": 0,
                "second": 0,
            },
            "status": {
                "slaMissFirst": 0,
                "slaMissSecond": 0
            }
        }


    dataLog("List of countries - " + str(listOfCountries))
    listOfCountries = listOfCountries

    outlook = client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    rootFolder = namespace.Folders[username]


    # print("Forecasting folder: " + forecasting_folder)

    # print("SA Update folder: " + sa_update_folder)

    # print("SA Report folder: " + sa_report_folder)


    dataLog("Folder found in Outlook: " + forecasting_folder)
    dataLog("Folder found in Outlook: " + sa_report_folder)
    dataLog("Folder found in Outlook: " + sa_update_folder)
    dataLog("Seconds settings: " + str(search_seconds) + " minutes")

    try:
        tasklist = pd.ExcelFile(path_to_task_list, engine='openpyxl')
        allNames = pd.read_excel(tasklist, "Schedulers")
        extraTasks = pd.read_excel(tasklist, "ExtraTasks")
        userByCountry = {}
        dataLog("Pairing user with country . . .")
        getSchedulerFromHeader = []
        for col in extraTasks.columns:
            getSchedulerFromHeader.append(col)
        FR_scheduler = getSchedulerFromHeader[1]
        userByCountry["FR"] = {"scheduler": FR_scheduler}
        userByCountry["DE"] = {"scheduler": extraTasks.iloc[0, 1]}
        userByCountry["ES"] = {"scheduler": extraTasks.iloc[0, 1]}
        userByCountry["IT"] = {"scheduler": FR_scheduler}
        

    except Exception as task_list_exception:
        dataLog("Something went wrong when getting data from the tasklist: " + str(task_list_exception))

    # ####### UPDATE DATASET WITH USERID'S

    for country in listOfCountries:
        dataSet[country]["user"] = userByCountry[country]["scheduler"]


    for i in range(0, len(allNames.index)):

        if allNames.iloc[i, 1] == userByCountry["DE"]["scheduler"]:
            userByCountry["DE"]["scheduler"] = allNames.iloc[i, 0]
            userByCountry["ES"]["scheduler"] = allNames.iloc[i, 0]

        elif allNames.iloc[i, 1] == userByCountry["IT"]["scheduler"]:
            userByCountry["IT"]["scheduler"] = allNames.iloc[i, 0]
            userByCountry["FR"]["scheduler"] = allNames.iloc[i, 0]

    dataLog("Scheduler ID and country pairing: DE - " + userByCountry["DE"]["scheduler"])
    dataLog("Scheduler ID and country pairing: ES - " + userByCountry["ES"]["scheduler"])
    dataLog("Scheduler ID and country pairing: IT - " + userByCountry["IT"]["scheduler"])
    dataLog("Scheduler ID and country pairing: FR - " + userByCountry["FR"]["scheduler"])


    # SEARCH FOR EMAILS AND RETURN A LIST OF MESSAGES

    def getEmails(folderName):  # THIS MUST BE CALLED THREE TIMES USING THE 3 DIFFERENT FOLDER NAMES

        timeCriteria = datetime.datetime.now() - dt.timedelta(minutes=120)
        list_of_emails = []
        try:
            emailFolder = rootFolder.Folders[folderName]

            messages = emailFolder.Items


            for message in messages:
                list_of_emails.append(message)

            dataLog("Retrieved emails from: " + folderName + " " + str(list_of_emails))

            return list_of_emails
        except Exception as email_exception:
            print(email_exception)
            dataLog(str(email_exception))
            return list_of_emails


    # UPDATE THE DATASET USING THE LISTS OF EMAILS

    def updateData(forecasting, saReports, saUpdates):
        fList = []
        uList = []
        saList = []
        try:
            for message in forecasting:

                fList.insert(0, message)
            for message in saUpdates:

                uList.insert(0, message)
            for message in saReports:

                saList.insert(0, message)
        except Exception as E:
            dataLog(str(E))
            print(E)

        # for each country look into the emails and find relevant data
        for country in listOfCountries:


            forecastString = "Estimated Expected Volumes - " + country + " - "
            saReportString = country + " stations - 24h Forecast " + backwordsToday
            saUpdateString = "[SA Update] " + country + " - " + slashToday

            # loop for forecasting emails
            try:
                for message in fList:
                    wave_time = int(str(message.CreationTime).split(" ")[1].split(":")[0])
                    # FILTER FIRST WAVES
                    today_formatted = ''.join(('- ', today1))
                    wave_t = wave_time
                    wave_first = firstWave
                    wave_second = secondWave
                    # if today in message.Subject and forecastString in message.Subject and wave_time <= firstWave and "RE: " not in message.Subject and "48h" not in message.Subject:
                    if today_formatted in message.Subject and forecastString in message.Subject and 'AM' in message.Subject and "RE: " not in message.Subject and "48h" not in message.Subject:
                    # if today in message.Subject or today_second_format in message.Subject:
                            # SEE IF EMAIL WAS ALREADY RECORDED
                        if message.CreationTime != dataSet[country]["forecast"]["received1"]:
                            dataSet[country]["forecast"]["received1"] = message.CreationTime
                            dataSet[country]["chimePing"]["first"] = 0

                        # FILTER SECOND WAVES
                    # elif today_second_format in message.Subject and forecastString in message.Subject and wave_time >= secondWave and wave_time <= 19 and "RE: " not in message.Subject and "48h" not in message.Subject:
                    elif today_formatted in message.Subject and forecastString in message.Subject and 'PM' in message.Subject and "RE: " not in message.Subject and "48h" not in message.Subject:
                        # SEE IF EMAIL WAS ALREADY RECORDED
                        if message.CreationTime != dataSet[country]["forecast"]["received2"]:
                            dataSet[country]["forecast"]["received2"] = message.CreationTime
                            dataSet[country]["chimePing"]["second"] = 0

            except Exception as E:
                dataLog(str(E))

            try:
                # LOOP FOR SAREPORT EMAILS
                for message in saList:
                    wave_time = int(str(message.CreationTime).split(" ")[1].split(":")[0])
                    # FILTER FIRST WAVES
                    if saReportString in message.Subject and wave_time <= firstWave and "RE: " not in message.Subject:

                        # SEE IF EMAIL WAS ALREADY RECORDED
                        if message.CreationTime != dataSet[country]["saReport"]["received1"]:
                            dataSet[country]["saReport"]["received1"] = message.CreationTime
                            dataSet[country]["chimePing"]["first"] = 0

                        # FILTER SECOND WAVES
                    elif saReportString in message.Subject and wave_time >= secondWave and wave_time <= 19 and "RE: " not in message.Subject:

                        # SEE IF EMAIL WAS ALREADY RECORDED
                        if message.CreationTime != dataSet[country]["saReport"]["received2"]:
                            dataSet[country]["saReport"]["received2"] = message.CreationTime
                            dataSet[country]["chimePing"]["second"] = 0

            except Exception as E:
                dataLog(str(E))
                print(E)

            # LOOP FOR SAUPDATE EMAILS
            try:
                for message in uList:
                    wave_time = int(str(message.CreationTime).split(" ")[1].split(":")[0])
                    if saUpdateString in message.Subject and wave_time <= firstWave and "RE: " not in message.Subject:

                        # SEE IF EMAIL WAS ALREADY RECORDED
                        if message.CreationTime != dataSet[country]["saUpdate"]["received1"]:
                            dataSet[country]["saUpdate"]["received1"] = message.CreationTime
                            dataSet[country]["chimePing"]["first"] = 0


                        # FILTER SECOND WAVES
                    elif saUpdateString in message.Subject and wave_time >= secondWave and wave_time <= 19 and "RE: " not in message.Subject:

                        # SEE IF EMAIL WAS ALREADY RECORDED
                        if message.CreationTime != dataSet[country]["saUpdate"]["received2"]:
                            dataSet[country]["saUpdate"]["received2"] = message.CreationTime
                            dataSet[country]["chimePing"]["second"] = 0


            except Exception as E:
                dataLog(str(E))
                print(E)
        dataLog("Updating dataset using timestamps of the emails. . .")
        dataLog(str(dataSet))
        test1 = dataSet
        return dataSet


    # FUNCTION THAT SENDS MESSAGES TO CHIME

    def broadcast_to_chime(url, message, user):
        userPing = "@" + user + " "
        result = False
        try:
            result = False
            session = requests.session()
            data = {"Content": userPing + message}
            params = {'format': 'application/json'}
            response = session.post(url, params=params, json=data)
            if response.status_code == 200:
                result = True

            return result
        except Exception as e:
            print("\nFailed to send Chime message: ", e)
            return result


    def broadcast_wave_to_chime(url, message):
        result = False
        try:
            result = False
            session = requests.session()
            data = {"Content": message}
            params = {'format': 'application/json'}
            response = session.post(url, params=params, json=data)
            if response.status_code == 200:
                result = True

            return result
        except Exception as e:
            print("\nFailed to send Chime message: ", e)
            return result


    # function that calculates SLA times
    def calculateSla(data):
        dataLog("Calculating SLA misses. . .")
        #print("SA Forecast First Wave, SA Forecast Second Wave, SA Update First Wave, SA Update Second Wave, SA Report First Wave, SA Report SecondWave")
        for country in listOfCountries:
            saReport1 = data[country]['saReport']['received1']
            saUpdate1 = data[country]['saUpdate']['received1']
            saReport2 = data[country]['saReport']['received2']
            saUpdate2 = data[country]['saUpdate']['received2']
            saForecast1 = data[country]['forecast']['received1']
            saForecast2 = data[country]['forecast']['received2']
            #print(saForecast1, saForecast2, saUpdate1, saUpdate2, saReport1, saReport2)

            if saReport1 != 0 and saForecast1 != 0:
                if saReport1 > saForecast1:
                    slaMissFirst = saReport1 - saForecast1
                    data[country]['status']['slaMissFirst'] = str(slaMissFirst).split(":")[0] + ":" + \
                                                            str(slaMissFirst).split(":")[1]
                else:
                    data[country]['status']['slaMissFirst'] = " Exception_detected"
                    dataLog("Exception detected for SLA for " + str(country))

            if saReport2 != 0 and saForecast2 != 0:
                if saReport2 > saForecast2:
                    slaMissSecond = saReport2 - saForecast2
                    data[country]['status']['slaMissSecond'] = str(slaMissSecond).split(":")[0] + ":" + \
                                                            str(slaMissSecond).split(":")[1]
                else:
                    data[country]['status']['slaMissSecond'] = " Exception_detected"
                    dataLog("Exception detected for SLA for " + str(country))

        dataLog("Calculate SLA function output: " + str(data))
        #pprint.pprint(data)
        return data


    def pingUser(data, url_list):
        dataLog("Checking what user needs to be pinged...")
        for country in listOfCountries:
            forecast1 = data[country]['forecast']['received1']
            forecast2 = data[country]['forecast']['received2']
            saReport1 = data[country]['saReport']['received1']
            saUpdate1 = data[country]['saUpdate']['received1']
            saReport2 = data[country]['saReport']['received2']
            saUpdate2 = data[country]['saUpdate']['received2']
            chimePing1 = data[country]['chimePing']['first']
            chimePing2 = data[country]['chimePing']['second']

            # check for FIRST WAVE

            if forecast1 != 0 and saReport1 == 0 and chimePing1 == 0:  # if report is NOT actioned but emails have been received and ping not sent to user
                broadcast_to_chime(url_list[0], "=> Forecast email for " + str(
                    country) + " has been received! Please check if the SA Logs have been updated and action the SA 24h Slope report.", data[country]["user"])
                dataLog(data[country]["user"] + " has been pinged on chime: => Forecast and SA Update email for " + str(
                    country) + " has been received! Please action the SA 24h Slope report.")
                data[country]['chimePing']['first'] = datetime.datetime.now()


            elif forecast1 != 0 and saReport1 == 0 and chimePing1 != 0:  # if report is NOT actioned but emails have been received and ping was already sent once for the user
                broadcast_to_chime(url_list[0], "=> Reminder => Forecast email for " + str(
                    country) + " has been received! Please check if the SA Logs have been updated and action the SA 24h Slope report.", data[country]["user"])
                data[country]['chimePing']['first'] = datetime.datetime.now()

            # check for SECOND WAVE

            if forecast2 != 0 and saReport2 == 0 and chimePing2 == 0:  # if report2 is NOT actioned but emails have been received and ping not sent to user
                broadcast_to_chime(url_list[0], "=> Forecast email for " + str(
                    country) + " has been received! Please check if the SA Logs have been updated and action the SA 24h Slope report.", data[country]["user"])

                data[country]['chimePing']['second'] = datetime.datetime.now()
            elif forecast2 != 0 and saReport2 == 0 and chimePing2 != 0:  # if report2 is NOT actioned but emails have been received and ping was already sent once for the user
                broadcast_to_chime(url_list[0], "=> Reminder => Forecast email for " + str(
                    country) + " has been received! Please check if the SA Logs have been updated and action the SA 24h Slope report.", data[country]["user"])

                data[country]['chimePing']['second'] = datetime.datetime.now()
        dataLog("Ping function output: " + str(data))
        return data


    def saveToExcel(data):
        myList = []
        headers = ["Country", "Scheduler", "First_Forecast", "First_Sa_Update", "First_Sa_Report", "Second_Forecast",
                "Second_Sa_Update", "Second_Sa_Report", "First_SLA", "Second_SLA"]
        for country in listOfCountries:
            user = str(data[country]['user'])
            forecast1 = str(data[country]['forecast']['received1'])
            forecast2 = str(data[country]['forecast']['received2'])
            saReport1 = str(data[country]['saReport']['received1'])
            saUpdate1 = str(data[country]['saUpdate']['received1'])
            saReport2 = str(data[country]['saReport']['received2'])
            saUpdate2 = str(data[country]['saUpdate']['received2'])
            firstSla = str(data[country]['status']['slaMissFirst']) + " (hh:mm)"
            secondSla = str(data[country]['status']['slaMissSecond']) + " (hh:mm)"
            myList.append(
                [country, user, forecast1, saUpdate1, saReport1, forecast2, saUpdate2, saReport2, firstSla, secondSla])

        dfObj = pd.DataFrame(data=myList, columns=headers)

        for x in range(dfObj.shape[0]):  # iterate over rows
            for j in range(dfObj.shape[1]):  # iterate over columns
                value = dfObj.iloc[x, j]  # get cell value
                if len(str(value)) > 17:
                    dfObj.iloc[x, j] = str(value).split(" ")[1].split(".")[0]
        
        os.chdir(r'\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\PM Shift\DHP1\SchedulingCompliance\\' + today + ' Compliance Uploads logs')
        with ExcelWriter(str(today) + ' SA Planning summary logs.xlsx') as writer:
            dfObj.to_excel(writer)
            dataLog("Saving data to Excel: \n\n" + str(dfObj))



            #print(" - - - > Run @ "+str(datetime.now()).split(".")[0] + " < - - - ")
            print ("\nSearching for Forecast Emails and SA Updates... \n")
            pprint.pprint(dfObj)
            return dfObj


    def waveReport(data, wave_rep, url_list):
        dataLog("Checking status of the wave report . . .")

        for country in listOfCountries:
            firstSla = data[country]['status']['slaMissFirst']
            secondSla = data[country]['status']['slaMissSecond']

            # 2021-05-01 10:55:18.722375
            # CHECK WHAT WAVES ARE FINISHED WITH THE SLA AND IF BOT HAS NOT SENT THE MESSAGE ALREADY
            if firstWaveTime < datetime.datetime.now() < firstWaveTimeEnd and wave_rep["first"] == 0:
                wave_rep["first"] = 1
                xlFile = pd.ExcelFile(str(today) + ' SA Planning summary logs.xlsx', engine='openpyxl')
                df = xlFile.parse(0)
                message1 = table_template[0] + table_template[1]
                for index, row in df.iterrows():
                    new_table_row = "|" + str(row["Country"]) + "|" + str(row["Scheduler"]) \
                                    + "|" + str(row["First_Forecast"]) + "|" + str(row["First_Sa_Update"]) + "|" + str(
                        row["First_Sa_Report"]) + "|" + str(row["First_SLA"]) \
                                    + "|" + str(row["Second_Forecast"]) + "|" + str(row["Second_Sa_Update"]) + "|" + str(
                        row["Second_Sa_Report"]) + "|" + str(row["Second_SLA"]) + "|\n"
                    message1 += new_table_row
                print("\nSending First Wave Report...")
                broadcast_wave_to_chime(url_list[1],
                                        message1 + "\n" + "|First Run| SLA Formula = SA_report time - SA_Forecast time")

                dataLog("Ping sent with First Wave Report!")
                print("\nFirst Wave Report has been send to the managers chime room.")
            else:
                dataLog("Finished checking first wave report")

            if secondWaveTime < datetime.datetime.now() < secondWaveTimeEnd and wave_rep["second"] == 0:
                wave_rep["second"] = 1
                xlFile = pd.ExcelFile(str(today) + ' SA Planning summary logs.xlsx', engine='openpyxl')
                df = xlFile.parse(0)
                message2 = table_template[0] + table_template[1]
                for index, row in df.iterrows():
                    new_table_row = "|" + str(row["Country"]) + "|" + str(row["Scheduler"]) \
                                    + "|" + str(row["First_Forecast"]) + "|" + str(row["First_Sa_Update"]) + "|" + str(row[
                                                                                                                        "First_Sa_Report"]) + "|" + str(
                        row["First_SLA"]) \
                                    + "|" + str(row["Second_Forecast"]) + "|" + str(row["Second_Sa_Update"]) + "|" + str(
                        row[
                            "Second_Sa_Report"]) + "|" + str(row["Second_SLA"]) + "|\n"
                    message2 += new_table_row

                print("\nSending Second Wave Report...")
                broadcast_wave_to_chime(url_list[1],
                                        message2 + "\n" + "|Second Run| SLA Formula = SA_report time - SA_Forecast time")
                dataLog("Ping sent with Second Wave Report!")
                print("\nSecond Wave Report has been send to the managers chime room.")
                
            else:
                dataLog("Finished checking second wave report")
            
        print("\nSUCCESS!: All SA Alerts have been already broadcasted")
        
        return wave_rep

    


    # GET ALL EMAILS
    forecastingEmails = getEmails(forecasting_folder)
    saReports = getEmails(sa_report_folder)
    saUpdates = getEmails(sa_update_folder)
    url_list = testing(are_we_testing)

    # UPDATE DICT WITH TIMESTAMPS FROM EMAILS
    updatedData = updateData(forecastingEmails, saReports, saUpdates)

    # CALCULATE SLA
    finalData = calculateSla(updatedData)

    # PING USERS ON CHIME
    pingUser(finalData, url_list)
    # SAVE DATA TO EXCEL
    saveToExcel(finalData)
    wave_rep = wave_report
    waveReport(finalData, wave_rep, url_list)

def SAPlanningAlertAutomated(today,currentdirectory):
    SAPlanningmain(today,currentdirectory)