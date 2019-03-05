import requests
import jiraConnector as jc
import pandas as pd
import common as cmn

import time
import datetime
import os
from seleniumrequests import Chrome
import urllib.parse as urlparse
from urllib.parse import urlencode

CHROMEDRIVER_PATH = 'C:\\Program Files\\Python36\\Scripts\\chromedriver.exe'
DOWNLOAD_PATH = cmn.DEFAULT_PATH + 'Downloads\\'
TIMESHEETS_PATH = cmn.DEFAULT_PATH + 'Documents\\Timesheets\\'

RUN_REPORT_URL = 'https://erp.netcracker.com/report/result.jsp?parent=7040459919013415288&report=7100462886013127576&object=3112170189013929210'
EXPORT_REPORT_URL = 'https://erp.netcracker.com/report/exportreport.jsp?expobject=7100462886013127576&adapterId=9135923224913544218'

PROJECTS_INFO = {'UUI_80' : {'id':'9146592119213395893', 'GP code' : 'NC.SDN.PRD.HOMUUIR8', 'start':'12/12/2016','end':'03/31/2017','projectLOE':1452}}
PROJECTS_SHORTLIST = ['UUI_80']

def getTimesheetFilename(project):
    return 'Timesheets_' + project + '_' + datetime.datetime.now().strftime("%Y-%m-%d") + '.xlsx'

def getTimesheetForPeriod(timesheet, start_date, end_date):
    timesheet['TS'] = timesheet.apply(lambda x : x['DAY'].timestamp(), axis=1)
    timesheet = timesheet.query('TS >= ' + str(start_date.timestamp()) + ' and ' + 'TS <= ' + str(end_date.timestamp()))
    return timesheet

#forcerun:
# - shortlist
# - all
# - one
# - no
def getTimesheetForProject(project, forcerun = 'shortlist'):
    filename = getTimesheetFilename(project)
    if filename in set(os.listdir(TIMESHEETS_PATH)):
        timesheet = pd.read_excel(TIMESHEETS_PATH + filename)
    else:
        print('Timesheets file not found, forcerun=' + forcerun)
        
        if forcerun == 'shortlist':
            list = PROJECTS_SHORTLIST
        elif forcerun == 'all':
            list = PROJECTS_INFO.keys()
        elif forcerun == 'one':
            list = [project]
        else:
            return None

        timesheet = makeReport(list)
        timesheet = timesheet[timesheet['PROJECT_ID'] == PROJECTS_INFO[project]['GP code']]

    return timesheet


def getURL(url, params):
    url_parts = list(urlparse.urlparse(url))
    query = dict(urlparse.parse_qsl(url_parts[4]))
    query.update(params)
    url_parts[4] = urlencode(query)

    return urlparse.urlunparse(url_parts)


def getReportRequest(URL, params, projects, start_date='', end_date=''):
    
    projects_ids = ','.join([PROJECTS_INFO[project]['id'] for project in projects])

    if start_date == '':
        start_date = "'" + min([datetime.datetime.strptime(PROJECTS_INFO[project]['start'], "%m/%d/%Y") for project in projects]).strftime("%m/%d/%Y") + "'"
    
    if end_date == '':
        end_date = "'" + datetime.datetime.now().strftime("%m/%d/%Y") + "'"

    return getURL(URL, {**{'param0':projects_ids
                                    ,'param1':start_date
                                    ,'param2':end_date}, **params})

def getRunReportRequest(projects, start_date='', end_date=''):
    
    runReportParams = {'param3':'NULL'
                    ,'param4':'NULL'
                    ,'param5':"'%'"
                    ,'param6':'NULL'}

    return getReportRequest(RUN_REPORT_URL, runReportParams, projects, start_date='', end_date='')


def getExportReportParams(projects, start_date='', end_date=''):

    projects_ids = ','.join([PROJECTS_INFO[project]['id'] for project in projects])

    if start_date == '':
        start_date = "'" + min([datetime.datetime.strptime(PROJECTS_INFO[project]['start'], "%m/%d/%Y") for project in projects]).strftime("%m/%d/%Y") + "'"
    
    if end_date == '':
        end_date = "'" + datetime.datetime.now().strftime("%m/%d/%Y") + "'"

    exportReportParams = {'param3':'NULL'
                        ,'param4':'NULL'
                        ,'report':'7100462886013127576'
                        ,'param5':"'%'"
                        ,'parent':'7040459919013415288'
                        ,'param6':'NULL'
                        ,'object':'3112170189013929210'}

    return {**{'param0':projects_ids
                ,'param1':start_date
                ,'param2':end_date}, **exportReportParams}

def getExportReportRequest(projects, start_date='', end_date=''):
    
    exportReportParams = {'param3':'NULL'
                        ,'param4':'NULL'
                        ,'report':'7100462886013127576'
                        ,'param5':"'%'"
                        ,'parent':'7040459919013415288'
                        ,'param6':'NULL'
                        ,'object':'3112170189013929210'}

    return getReportRequest(EXPORT_REPORT_URL, exportReportParams, projects, start_date='', end_date='')


def runReport(projects, start_date='', end_date=''):
    
    driver = Chrome(CHROMEDRIVER_PATH)  # Optional argument, if not specified will search path.
    #driver.set_window_size(300,200)
    
    runReportRequest = getRunReportRequest(projects, start_date, end_date)
    driver.get(runReportRequest)
    
    return driver


def exportReport(projects, driver=None, start_date='', end_date=''):

    if driver is None:
        driver = Chrome(CHROMEDRIVER_PATH)  # Optional argument, if not specified will search path.
        #driver.set_window_size(100,100)
    
    exportReportRequest = getExportReportRequest(projects, start_date, end_date)

    before = os.listdir(DOWNLOAD_PATH)
    
    print("Now export to Excel 2007 manually and press Enter")
    input()

    #res = driver.request('POST', exportReportRequest, data)

    #time.sleep(10)

    loaded = False
    attempts = 20

    while not loaded and attempts > 0:
        after = os.listdir(DOWNLOAD_PATH)
        change = set(after) - set(before)
        if len(change) == 1:
            file_name = change.pop()
            if file_name.find('crdowndload') < 0:
                loaded = True
                driver.quit()
                return file_name
        print("Waiting for download " + str(attempts))
        time.sleep(5)
        attempts = attempts - 1

    return ''

#If you need standalone Timesheet export, please use this function
#Example: ts = ter.makeReport(['ESO_NEW_ARCH','ESO2_R1'])
def makeReport(projects, start_date='', end_date='', path = TIMESHEETS_PATH):

    driver = runReport(projects)
    time.sleep(10)
    TERfile = exportReport(projects,driver=driver)

    if TERfile == '':
        print('Report export FAILED')
        return None

    timesheet = pd.read_excel(DOWNLOAD_PATH + TERfile,
                                header=0,
                                skiprows=12+len(projects),
                                parse_cols='A:U')

    timesheet.dropna(subset=['EMPLOYEE NAME'],inplace=True,how='all')

    timesheet['MD'] = timesheet['HOUR']/8

    columns=['EMPLOYEE NAME',
             'TASK GROUP',
             'TASK',
             'WORK ITEM ID',
             'WORK ITEM',
             'DAY',
             'WEEK',
             'HOUR',
             'MD',
             'COMMENT']

    for project in projects:
        filename = getTimesheetFilename(project)
        timesheet_for_project = getTimesheetForPeriod(timesheet[timesheet['PROJECT_ID'] == PROJECTS_INFO[project]['GP code']],
                                                    datetime.datetime.strptime(PROJECTS_INFO[project]['start'], "%m/%d/%Y"),
                                                    datetime.datetime.strptime(PROJECTS_INFO[project]['end'], "%m/%d/%Y"))
        timesheet_for_project.to_excel(path + filename,
                                        merge_cells=False,
                                        columns = columns,
                                        index=False)

    return timesheet


'''

TERfile = 'Timesheets by project-08-14-2017 11-07.xlsx'
projects = ['ES20_R2']
project = projects[0]
timesheet[timesheet['DAY'] >= datetime.datetime.strptime(PROJECTS_INFO[project]['start'], "%m/%d/%Y") and timesheet['DAY'] <= datetime.datetime.strptime(PROJECTS_INFO[project]['end'], "%m/%d/%Y")]
datetime.datetime.strptime(PROJECTS_INFO[project]['start'], "%m/%d/%Y")

timesheet[timesheet['TASK GROUP'] == 'DT Intraselect POC'].to_excel('C:\\Users\\rudu0916\\Documents\\Timeshhets_VDOC_R3_DT Intraselect POC.xlsx',
                        merge_cells=False,columns = columns,index=False)


timesheet.columns

timesheet[['TASK GROUP','EMPLOYEE NAME','HOUR']].groupby(by=['TASK GROUP','EMPLOYEE NAME']).sum()
timesheet[['TASK GROUP','EMPLOYEE NAME','HOUR']]

timesheet.groupby(by='EMPLOYEE NAME').sum()


driver = webdriver.Chrome(CHROMEDRIVER_PATH)  # Optional argument, if not specified will search path.
driver.set_window_size(100,100)
driver.get(runReportRequest)

#time.sleep(30) # Let the user actually see something!

before = os.listdir(DOWNLOAD_PATH)

driver.get(exportReportRequest)

time.sleep(30)

after = os.listdir(DOWNLOAD_PATH)
change = set(after) - set(before)
if len(change) == 1:
    file_name = change.pop()
else:
    print("More than one file or no file downloaded")

driver.quit()

TERfile = file_name

a = {'local'+
        'storage'}

'''