import jiraConnector as jc
import common as cmn
import pandas as pd
import re
import datetime as dt
import numpy as np
import pytz
import time
import tzlocal
from collections import namedtuple
import warnings
from xlsxwriter.utility import xl_rowcol_to_cell
import pandas.io.formats.excel
import numbers
import WBStools as wbs
import TERtools as ter
import dailyreport as dr

#HINT: sprintRow = <row number in Excel> - 3
SPRINTS_INFO = {'VDOCSMR6' : {'phases':['DEV'], 'sprintRow': 8,'parse_cols': 'A:M','split_teams': False,
                            'RPpath': cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\AVP\\R6\\NC.SDN.PRD.VDC.R6.SM_RP_V0_01.29.2018.xlsm'}
               ,'VDOCSMR7' : {'phases':['DEV'], 'sprintRow': 8,'parse_cols': 'A:M','split_teams': False,
                            'RPpath': cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\AVP\\R7\\NC.SDN.PRD.VDC.R6.SM_RP_V1_03.20.2018 proposed.xlsm'}
               ,'VDOCDPR6' : {'phases':['DEV'], 'sprintRow': 7,'parse_cols': 'A:M','split_teams': False,
                            'RPpath': cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\AVP\\R6\\NC.SDN.PRD.VDC.R6.DP_RP_V0_02.19.2018.xlsm'}
               ,'VDOCDPR7' : {'phases':['DEV'], 'sprintRow': 7,'parse_cols': 'A:M','split_teams': False,
                            'RPpath': cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\AVP\\R6\\NC.SDN.PRD.VDC.R6.DP_RP_V0_02.19.2018.xlsm'}
               ,'ES20_R3': {'phases':['R3'], 'sprintRow': 36,'parse_cols': 'A:D,AA:AO','split_teams': False,
                            'RPpath': cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\ES2.0\\NC.SDN.ES2.0_RP_V4_09.28.2017.xlsm'}
               ,'ES20_R4': {'phases':['R4'], 'sprintRow': 60,'parse_cols': 'A:D,AO:BA','split_teams': False,
                            'RPpath': cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\ES2.0\\NC.SDN.ES2.0_RP_V11_12.11.2017_R5.xlsm'}
               ,'CM91': {'phases':['DEV'], 'sprintRow': 9,'parse_cols': 'A:M','split_teams': False,
                            'RPpath': cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\HOM\\CM\\NC.SDN.PRD.HOMCM91_RP_V6_10.17.2017.xlsm'}
               ,'CM911': {'phases':['DEV'], 'sprintRow': 9,'parse_cols': 'A:L','split_teams': False,
                            'RPpath': cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\HOM\\CM\\NC.SDN.PRD.HOMCM911_RP_V0_11.28.2017.xlsm'}
               ,'CM92': {'phases':['DEV','BA'], 'sprintRow': 9,'parse_cols': 'A:Q','split_teams': True,'unified_sprints': True,
                            'RPpath': cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\HOM\\CM\\NC.SDN.PROD.HOMCM92_RP_V5_02.26.2018.xlsm'}
               ,'CM93': {'phases':['DEV','BA'], 'sprintRow': 8,'parse_cols': 'A:W','split_teams': False,'unified_sprints': True,
                            'RPpath': cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\HOM\\CM\\NC.SDN.PROD.HOMCM93_RP_V5_07.18.2018 V2.xlsm'}
               ,'ESO91': {'phases':['DEV','QA','BA'], 'sprintRow': 11,'parse_cols': 'A:V','split_teams': True,'unified_sprints': True,
                            'RPpath': cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\ESO\\NC.SDN.PRD.HOMSO91_RP_V15_12.18.2017.xlsm'}}

def getKPI(a, b, c):

    x = 1 - a
    y = 1 - b
    z = 1 - c

    return round(x*x*2 + y/5 + z*z*5, 3)

GREEN_THRESHOLD = getKPI(0.6, 0.9, 0.9)
YELLOW_THRESHOLD = getKPI(0.5, 0.7, 0.7)

def getTeam(sprint_name, board_name = ''):
    if board_name == '':
        m = re.search('((\W*\w+)*).*(?=[Ss]p)',sprint_name)
        if not (m is None):
            return m.group(1)
        else:
            return sprint_name
    else:   
        m = re.findall('\[(.+)\]',board_name)
        if len(m) == 0:
            return board_name
        else:
            return m[0]

def getSprint(sprint_name):
    m = re.search('([Ss]p[\\w\\s]*[\\d]+).*',sprint_name)\
    #m = re.search('[Ss]p.+',sprint_name)
    if not (m is None):
        #return m.group(0)
        return m.group(1)
    else:
        return sprint_name

def trySprint(id, jira):
    #warnings.filterwarnings('error')
    try:
        return jira.sprint(id)
    except Exception as e:
        print(e.text)
        mySprint = namedtuple('mySprint', 'name startDate endDate state completeDate')
        return mySprint(name='N/A', startDate='N/A', endDate='N/A', state='N/A', completeDate='N/A')

def getSprints(jira, board_id, project):
    raw_sprints = jira.sprints(board_id)
    board_name = ''

    if 'unified_sprints' in SPRINTS_INFO[project]:
        if SPRINTS_INFO[project]['unified_sprints']:
            for board in jira.boards():
                if board.id == board_id:
                    board_name = board.name
                    break
    
    sprints_keys = list([x.id for x in raw_sprints])
    sprints_list = list(map(lambda x: trySprint(x, jira), sprints_keys))
    sprints_list = list(map(lambda x:
                        [board_id,
                         getTeam(x.name, board_name),
                         getSprint(x.name),
                         x.name,
                         x.startDate,
                         x.endDate,
                         x.state,
                         x.completeDate],
                        sprints_list))

    sprints_df = pd.DataFrame(sprints_list,
                            index=sprints_keys,
                            columns=['Board',
                                     'Team',
                                    'Sprint',
                                    'Sprint_name',
                                    'startDate',
                                    'endDate',
                                    'state',
                                    'completeDate'])
    return sprints_df

def getScopeChangeBurndownChart(jira, board_id, sprint_id):
    return jira._get_json(
                'rapid/charts/scopechangeburndownchart?rapidViewId=%s&sprintId=%s'
                % (board_id, sprint_id),
                base=jira.AGILE_BASE_URL)

def parseBurndownTimestamp(ts):
    if not isinstance(ts, numbers.Number):
        tsi = int(ts)
    else:
        tsi = ts
    localzone = tzlocal.get_localzone()
    naive = dt.datetime.fromtimestamp(tsi / 1000, tz = pytz.utc).replace(tzinfo = None)
    return localzone.localize(naive)

def getCurrentTimeFromBurndown(scopeChangeBurndownChart):
    return parseBurndownTimestamp(scopeChangeBurndownChart['now'])

#TODO account for different sprints, not last
def getSprintIDs(sprints, sprint=0):
    return sprints.query('state == "ACTIVE"')

def parseBurndown(Burndown):

    columns=['timestamp_str',
             'timeSpent',
            'oldEstimate',
            'newEstimate',
            'deltaEstimate',
            'changeDate',
            'changeDate_str',
            'notDone',
            'done',
            'newStatus',
            'added',
            'initialScope',
            'additionalScope',
            'descope']

    data = []
    index1 = []
    index2 = []
    index3 = []
    
    startTime = Burndown['startTime']
    endTime = Burndown['endTime']

    if 'completeTime' in Burndown:
        completeTime = Burndown['completeTime']
    else:
        completeTime = None

    for timestamp, changes in Burndown['changes'].items():
        timestamp_str = parseBurndownTimestamp(timestamp).strftime('%Y.%m.%d %H:%M')
        timestamp = int(timestamp)
        for change in changes:
            key = change['key']
            notDone = ''
            done = ''
            newStatus = ''
            timeSpent = 0
            oldEstimate = None
            newEstimate = None
            deltaEstimate = 0
            changeDate = None
            changeDate_str = ''
            added = ''
            initialScope = ''
            additionalScope = ''
            descope = ''

            if key in Burndown['issueToParentKeys'].values():
                index2 = index2 + [key]
            elif key in Burndown['issueToParentKeys'].keys():
                index2 = index2 + [Burndown['issueToParentKeys'][key]]
            else:
                index2 = index2 + ['']

            if 'column' in change:
                column = change['column']

                if 'notDone' in column:
                    notDone = column['notDone']

                if 'done' in column:
                    done = column['done']

                if 'newStatus' in column:
                    newStatus = column['newStatus']

            if 'timeC' in change:
                timeC = change['timeC']

                if 'timeSpent' in timeC:
                    timeSpent = timeC['timeSpent']/3600

                if 'oldEstimate' in timeC:
                    oldEstimate = timeC['oldEstimate']/3600
                else:
                    oldEstimate = 0

                if 'newEstimate' in timeC:
                    newEstimate = timeC['newEstimate']/3600
                else:
                    newEstimate = 0

                deltaEstimate = newEstimate - oldEstimate
                    
                if 'changeDate' in timeC:
                    changeDate = timeC['changeDate']
                    changeDate_str = parseBurndownTimestamp(timeC['changeDate']).strftime('%Y.%m.%d %H:%M')
                    changeDate = int(changeDate)

            if 'added' in change:
                added = change['added']
                if added:
                    if timestamp <= startTime:
                        initialScope = True
                    else:
                        additionalScope = True
                else:
                    if (timestamp > startTime) and (timestamp < cmn.nvl(completeTime,endTime)):
                        descope = True
                    elif timestamp <= startTime:
                        raise ValueError('Added is False before sprint start')
     
            data = data + [[timestamp_str,
                            timeSpent,
                           oldEstimate,
                           newEstimate,
                           deltaEstimate,
                           changeDate,
                           changeDate_str,
                           notDone,
                           done,
                           newStatus,
                           added,      
                           initialScope,
                           additionalScope,
                           descope]]
            index1 = index1 + [timestamp]
            index3 = index3 + [key]
    res = pd.DataFrame(data, index=[index1, index2, index3], columns = columns)
    res.index.levels[0].name = 'timestamp'
    res.index.levels[1].name = 'parent'
    res.index.levels[2].name = 'key'
    return res

def getInitialEstimate(parsedBurndown, key, startTime):
    fromTime = max(startTime, parsedBurndown.query('key == "' + key + '" and added == True').iloc[0].name[0])
    filteredBurndown = parsedBurndown.query('timestamp <= ' + str(fromTime) +
                                            ' and key == "' + key + '"')
    #return pd.Series({'initialEstimate' : filteredBurndown['deltaEstimate'].sum()})
    return filteredBurndown['deltaEstimate'].sum()

def getTimeSpent(parsedBurndown, key, startTime, endTime):
    fromTime = max(startTime, parsedBurndown.query('key == "' + key + '" and added == True').iloc[0].name[0])
    descoped = parsedBurndown.query('key == "' + key + '" and added == False')
    if len(descoped) > 0:
        toTime = min(endTime, descoped.iloc[0].name[0])
    else:
        toTime = endTime
    filteredBurndown = parsedBurndown.query('timestamp >= ' + str(startTime) +
                                            ' and timestamp <= ' + str(toTime) +
                                            ' and key == "' + key + '"')
    #return pd.Series({'totalTimeSpent' : filteredBurndown['timeSpent'].sum()})
    return filteredBurndown['timeSpent'].sum()

def getRemainingEstimate(parsedBurndown, key, endTime):
    filteredBurndown = parsedBurndown.query('timestamp <= ' + str(endTime) +
                                            ' and key == "' + key + '"')
    return filteredBurndown['deltaEstimate'].sum()

def getCompletedDate(parsedBurndown, key, endTime):
    filteredBurndownDone = parsedBurndown.query('timestamp <= ' + str(endTime) +
                                                ' and done == True '
                                                ' and key == "' + key + '"')
    if len(filteredBurndownDone) == 0:
        return None
    lastDone = max(filteredBurndownDone.index.get_level_values(0))
    filteredBurndownNotDone = parsedBurndown.query('timestamp >= ' + str(lastDone) +
                                                    ' and timestamp <= ' + str(endTime) +
                                                    ' and notDone == True '
                                                    ' and key == "' + key + '"')
    if len(filteredBurndownNotDone) == 0:
        return lastDone
    else:
        return None

def issueDetails(issue, parsedBurndown, sprintIssues, now, startTime, endTime, jira, getdailyrep):
    parent = issue.name[1]
    key = issue.name[2]
    initialScope = issue['initialScope']
    additionalScope = issue['additionalScope']
    decsopeFilter = parsedBurndown.query('key == "' + key + '" and added == False')
    #TODO doesn't account for case when issue removed from sprint and added again during sprint
    if len(decsopeFilter) > 0:
        descope = True
        descoped = decsopeFilter.iloc[0].name[0]
        descoped_str = decsopeFilter.iloc[0]['timestamp_str']
    else:
        descope = ''
        descoped = None
        descoped_str = ''
    added = issue.name[0]
    added_str = issue['timestamp_str']
    completed = getCompletedDate(parsedBurndown, key, endTime)
    if completed is None:
        completed_str = ''
        remainingEstimate = getRemainingEstimate(parsedBurndown, key, min(now, endTime, cmn.nvl(descoped,endTime)))
    else:
        completed_str = parseBurndownTimestamp(completed).strftime('%Y.%m.%d %H:%M')
        remainingEstimate = 0
    initialEstimate = getInitialEstimate(parsedBurndown, key, startTime)
    timeSpent = getTimeSpent(parsedBurndown, key, startTime, cmn.nvl(descoped,endTime))
    jiraIssue = [x for x in sprintIssues if x.key == key][0]
    Summary = jiraIssue.fields.summary
    Status = jiraIssue.fields.status.name
    Status_mapped = jc.statuses_mapped.loc[jiraIssue.fields.status.name]['Type']
    assignee = jc.get_assignee(jiraIssue.fields.assignee)
    if getdailyrep:
        dailyreport = dr.dailyreport(jira,jiraIssue)
    else:
        dailyreport = ""
    
    return pd.Series({'parent' : parent,
                      'key' : key,
                      'initialScope' : initialScope,
                      'additionalScope' : additionalScope,
                      'descope' : descope,
                      'added' : added,
                      'added_str' : added_str,
                      'descoped' : descoped_str,
                      'completed' : completed,
                      'completed_str' : completed_str,
                      'initialEstimate' : initialEstimate,
                      'timeSpent' : timeSpent,
                      'remainingEstimate' : remainingEstimate,
                      'Summary' : Summary,
                      'Status' : Status,
                      'Assignee' : assignee,
                      'Status_mapped' : Status_mapped,
                      'dailyreport' : dailyreport})

def detailedSprintReport(jira, parsedBurndown, Burndown, getdailyrep):
    now = Burndown['now']
    startTime = Burndown['startTime']
    endTime = Burndown['endTime']
    if 'completeTime' in Burndown:
        completeTime = Burndown['completeTime']
    else:
        completeTime = None

    sprintScope = parsedBurndown.query('added == True')

    if len(sprintScope) == 0:
        return None

    jql = 'issue in (' + ','.join(str(key) for key in sprintScope.index.get_level_values(2)) + ')'
    sprintIssues = jira.search_issues(jql,maxResults=1000)

    print('Preparing detailed report...')

    sprintDetailedReport =  sprintScope.apply(lambda x: issueDetails(x, parsedBurndown, sprintIssues, now, startTime, cmn.nvl(completeTime,endTime), jira, getdailyrep),axis=1)
    sprintDetailedReport = sprintDetailedReport.query('completed != completed or completed > added')
    sprintDetailedReport['notParent'] = sprintDetailedReport.apply(lambda row: row.name[1] != row.name[2],axis=1)
    sprintDetailedReport['Parent'] = sprintDetailedReport.index.get_level_values(1)
    sprintDetailedReport['Key'] = sprintDetailedReport.index.get_level_values(2)
    sprintDetailedReport = sprintDetailedReport.sort_values(['Parent','notParent','Key'])
    #sprintDetailedReport.to_excel('C:\\Users\\rudu0916\\Documents\\SprintReport_IPAM_Sp8.xlsx')

    #Initial scope
    #initialScope = parsedBurndown.query('initialScope == True')
    #initialScope = initialScope.join(initialScope.apply(lambda x: getInitialEstimate(parsedBurndown, x.name[2], startTime),axis=1))
    #initialScope = initialScope.join(initialScope.apply(lambda x: getTimeSpent(parsedBurndown, x.name[2], startTime, cmn.nvl(completeTime,endTime)),axis=1))
    #Additional scope
    #additionalScope = parsedBurndown.query('additionalScope == True')
    #additionalScope = additionalScope.join(additionalScope.apply(lambda x: getInitialEstimate(parsedBurndown, x.name[2],startTime),axis=1))
    #additionalScope = additionalScope.join(additionalScope.apply(lambda x: getTimeSpent(parsedBurndown, x.name[2], x.name[0], cmn.nvl(completeTime,endTime)),axis=1))
    #Descoped
    #descoped = parsedBurndown.query('descope == True')
    #descoped = descoped.join(descoped.apply(lambda x: getInitialEstimate(parsedBurndown, x.name[2], startTime),axis=1))
    #descoped = descoped.join(additionalScope.apply(lambda x: getTimeSpent(parsedBurndown, x.name[2], startTime, cmn.nvl(completeTime,endTime)),axis=1))

    return sprintDetailedReport

def getPerformance(project, sprint, sprintDetailedReport, sprints_dates, assignee=None):

    Capacity = getCapacity(project, sprint, sprints_dates, assignee)
    RemainingCapacity = getRemainingCapacity(Capacity, sprint, sprints_dates)
    TERs = getTERs(project, sprint, sprints_dates, assignee)
    Velocity = 0
    Efficiency = 0

    sprintDetailedReport_filtered = sprintDetailedReport.query('notParent==True')
    
    if not assignee is None:
        sprintDetailedReport_filtered = sprintDetailedReport_filtered.query('Assignee=="' + assignee + '"')
    
    PlannedTasksTotal = len(sprintDetailedReport_filtered)

    if PlannedTasksTotal == 0:
        PlannedTotal = 0
        Fact = 0
        RemainingTotal = 0
        completedIssues = 0
        currCompletedIssues = 0
        ImplementedRate = 0
        CurrImplementedRate = 0
    else:
        PlannedTotal = sprintDetailedReport_filtered['initialEstimate'].sum()
        Fact = sprintDetailedReport_filtered['timeSpent'].sum()
        RemainingTotal = sprintDetailedReport_filtered['remainingEstimate'].sum()
        completedIssues = sprintDetailedReport_filtered.query('completed_str != ""')
        currCompletedIssues = sprintDetailedReport_filtered.query('Status_mapped == "In QA" or Status_mapped == "Done"')
        ImplementedRate = getImplementedRate(sprintDetailedReport_filtered, completedIssues)
        CurrImplementedRate = getImplementedRate(sprintDetailedReport_filtered, currCompletedIssues)

    if PlannedTotal > 0:
        FactPlannedRate = Fact / PlannedTotal
    else:
        FactPlannedRate = 0

    if Capacity > 0:
        PlannedCapacityRate = PlannedTotal / Capacity
    else:
        PlannedCapacityRate = 0

    if PlannedTasksTotal > 0:
        InitiallyPlannedTasks = len(sprintDetailedReport_filtered.query('initialScope == True'))
        AdditionalScope = len(sprintDetailedReport_filtered.query('additionalScope == True'))
        DescopedTasks = len(sprintDetailedReport_filtered.query('descope == True'))
        CompletedTasks = len(completedIssues)
        CompletedTasksRate = CompletedTasks / PlannedTasksTotal
        CurrCompletedTasksRate = len(currCompletedIssues) / PlannedTasksTotal
        KPI = getKPI(PlannedCapacityRate, FactPlannedRate, ImplementedRate)


        if TERs > 0:
            Velocity = Fact / TERs
            Efficiency = ImplementedRate * PlannedTotal / TERs

    else:
        InitiallyPlannedTasks = 0
        AdditionalScope = 0
        DescopedTasks = 0
        CompletedTasks = 0
        CompletedTasksRate = 0
        CurrCompletedTasksRate = 0
        KPI = 99

    return pd.Series({'Capacity' : Capacity,
                      'RemainingCapacity' : RemainingCapacity,
                      'PlannedTotal' : PlannedTotal,
                      'Fact' : Fact,
                      'RemainingTotal' : RemainingTotal,
                      'TERs' : TERs,
                      'ImplementedRate' : ImplementedRate,
                      'CurrImplementedRate' : CurrImplementedRate,
                      'FactPlannedRate' : FactPlannedRate,
                      'PlannedCapacityRate' : PlannedCapacityRate,
                      'InitiallyPlannedTasks' : InitiallyPlannedTasks,
                      'AdditionalScope' : AdditionalScope,
                      'DescopedTasks' : DescopedTasks,
                      'PlannedTasksTotal' : PlannedTasksTotal,
                      'CompletedTasks' : CompletedTasks,
                      'CompletedTasksRate' : CompletedTasksRate,
                      'CurrCompletedTasksRate' : CurrCompletedTasksRate,
                      'Velocity' : Velocity,
                      'Efficiency' : Efficiency,
                      'KPI' : KPI})

def teamSprintReport(project, sprint, sprints_dates, jira, sprintDetailedReport):

    #TODO account for case when RP is not present
    if sprints_dates is None:
        return 0

    RP = getFilteredRP(project, sprint, sprints_dates)
    
    sprint_name = sprint['Sprint']

    RP = RP[sprint_name].to_frame()
    RP.dropna(inplace = True)

    sprintTeamReport = sprintDetailedReport.query('notParent==True')
    sprintTeamReport = sprintTeamReport[['Assignee','key']].groupby('Assignee').count()

    sprintTeamReport = sprintTeamReport.join(RP, how = 'outer')

    sprintTeamReport = sprintTeamReport.join(
        sprintTeamReport.apply(lambda row: getPerformance(project,
                                                          sprint,
                                                          sprintDetailedReport,
                                                          sprints_dates,
                                                          assignee = row.name),axis=1))

    return sprintTeamReport.sort_values('KPI')


def getFilteredRP(project, jira_sprint, sprints_dates):

    RP = wbs.loadRP(SPRINTS_INFO[project]['RPpath'],
                    index_col=3,
                    parse_cols=SPRINTS_INFO[project]['parse_cols'])

    RP = RP[pd.notnull(RP.index)]

    for (sprint, params) in sprints_dates.items():
        RP[sprint] = RP[list(range(params['start_id'],params['end_id']+1))].apply(sum,axis=1)

    RP = RP.query('Phase in ' + str(SPRINTS_INFO[project]['phases']))
    
    if SPRINTS_INFO[project]['split_teams']:
        Team = jira_sprint['Team']
        RP['TeamMatch'] = RP.apply(lambda x : Team.find(x['Team']),axis=1)
        RP = RP.query('TeamMatch >= 0')

    return RP

def getSprintsDates(project):

    if project not in SPRINTS_INFO:
        print('No sprints info for project ' + project)
        return None

    sprints_dates = wbs.getSprints(SPRINTS_INFO[project]['RPpath'],
                            parse_cols = SPRINTS_INFO[project]['parse_cols'],
                            sprintRow = SPRINTS_INFO[project]['sprintRow'])

    return sprints_dates

def getCapacity(project, sprint, sprints_dates, assignee=None):

    if sprints_dates is None:
        return 0

    RP = getFilteredRP(project, sprint, sprints_dates)
    
    if not assignee is None:
        RP = RP[RP.index == assignee]

    sprint_name = sprint['Sprint']

    return RP[sprint_name].sum()

def getRemainingCapacity(capacity, sprint, sprints_dates):

    if sprints_dates is None:
        return 0

    sprint_name = sprint['Sprint']

    sprint_len = np.busday_count(sprints_dates[sprint_name]['start_date'], sprints_dates[sprint_name]['end_date'])
    remaining_days = np.busday_count(dt.datetime.now(), sprints_dates[sprint_name]['end_date'])

    if remaining_days > 0 and sprint_len > 0:
        return capacity * remaining_days/sprint_len
    else:
        return 0

def getTERs(project, sprint, sprints_dates, assignee=None):

    if sprints_dates is None:
        return 0

    timesheet = ter.getTimesheetForProject(project)

    sprint_name = sprint['Sprint']

    if not assignee is None:
        timesheet = timesheet[timesheet['EMPLOYEE NAME'] == assignee]
    else:
        RP = getFilteredRP(project, sprint, sprints_dates)        
        RP = RP[RP[sprint_name] > 0]

        timesheet = timesheet[timesheet['EMPLOYEE NAME'].isin(RP.index.values)]

    if len(timesheet) == 0:

        msg = 'No TER entries are found'
        
        if not assignee is None:
            msg = msg + ' for ' + assignee

        print(msg)

        return 0

    timesheet = ter.getTimesheetForPeriod(timesheet, sprints_dates[sprint_name]['start_date'], sprints_dates[sprint_name]['end_date'])

    return timesheet['HOUR'].sum()

def estimate(issue):
    return max(issue['initialEstimate'], issue['timeSpent'] + issue['remainingEstimate'])

def getImplementedRate(sprintDetailedReport, completedIssues):
    #TODO maybe need to exclude descoped issues
    
    if len(completedIssues) > 0:
        implemented_estimate = sum(completedIssues.apply(estimate,axis=1))
    else:
        return 0
    
    total_estimate = sum(sprintDetailedReport.apply(estimate,axis=1))

    if total_estimate > 0:
        return implemented_estimate / total_estimate
    else:
        return 0

def aggregatedSprintReport(project, sprint, sprintDetailedReport, Burndown, sprints_dates):
    Team = sprint['Team']
    Sprint = sprint['Sprint']
    state = sprint['state']

    startTime = parseBurndownTimestamp(Burndown['startTime']).strftime('%Y.%m.%d %H:%M')
    endTime = parseBurndownTimestamp(Burndown['endTime']).strftime('%Y.%m.%d %H:%M')
    
    if 'completeTime' in Burndown:
        completeTime = parseBurndownTimestamp(Burndown['completeTime']).strftime('%Y.%m.%d %H:%M')
    else:
        completeTime = ''

    sprintAggregatedReport = pd.Series({'Team' : Team,
                                        'Sprint' : Sprint,
                                        'startTime' : startTime,
                                        'endTime' : endTime,
                                        'state' : state,
                                        'completeTime' : completeTime})

    performance = getPerformance(project, sprint, sprintDetailedReport, sprints_dates)

    sprintAggregatedReport = pd.concat([sprintAggregatedReport, performance])

    return sprintAggregatedReport

def format_DetailedReport(workbook,
                          worksheet,
                          number_rows,
                          sprintAggregatedReport,
                          Aggregated_columns,
                          Aggregated_labels):

    format_all = workbook.add_format(cmn.format_all)
    header_fmt = workbook.add_format(cmn.header_fmt)
    effort_fmt = workbook.add_format(cmn.effort_fmt)
    percent_fmt = workbook.add_format(cmn.percent_fmt)

    worksheet.set_column('A:B', 12, format_all)
    worksheet.set_column('C:C', 55, format_all)
    worksheet.set_column('D:D', 20, format_all)
    worksheet.set_column('E:F', 9, format_all)
    worksheet.set_column('G:I', 8, effort_fmt)
    worksheet.set_column('J:K', 12, format_all)
    worksheet.set_column('L:N', 8, format_all)
    worksheet.set_column('O:O', 12, format_all)
    worksheet.set_column('P:P', 8, format_all)
    worksheet.set_column('Q:Q', 55, format_all)

    worksheet.set_row(0, None, header_fmt)
    
    
    format_all_bolditalic = workbook.add_format({**cmn.format_all, **{'bold' : True, 'italic' : True}})

    format_alert = workbook.add_format(cmn.format_alert)
    format_warn = workbook.add_format(cmn.format_warn)
    
    format_inprogress = workbook.add_format(cmn.format_inprogress)
    format_impl = workbook.add_format(cmn.format_impl)
    format_cancelled = workbook.add_format(cmn.format_cancelled)

    #Issue is parent
    worksheet.conditional_format('A2:F{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=$A2=$B2', 'format': format_all_bolditalic})    
    
    #Zero Remaining for incompleted task
    worksheet.conditional_format('I2:I{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=AND(OR($F2="To Do",$F2="In Progress"),$I2=0,$A2<>$B2)', 'format': format_warn})
    #Issue is in progress
    worksheet.conditional_format('A2:P{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=$F2="In Progress"', 'format': format_inprogress})
    #Issue is completed in sprint
    worksheet.conditional_format('A2:P{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=$K2<>""', 'format': format_impl})
    #Issue status is In QA/Done and issue is not identified as completed in sprint
    worksheet.conditional_format('E2:F{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=AND(OR($F2="In QA",$F2="Done"),$K2="")', 'format': format_impl})
    #Cancelled or descoped
    worksheet.conditional_format('A2:P{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=OR($F2="Cancelled",$N2=TRUE)', 'format': format_cancelled})

    #Progress bars
    worksheet.conditional_format('K{}:K{}'.format(number_rows+3,number_rows+6), cmn.format_bar_impl)

    #Spent effort exceeds estimate for over then 20%
    worksheet.conditional_format('H2:H{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=($H2/$G2)>=1.2', 'format': format_alert})
    
    #Gained effort exceeds gained progress for over then 20%
    worksheet.conditional_format('H{}:H{}'.format(number_rows+4,number_rows+4),
                                 {'type': 'formula',
                                  'criteria': '=($H{}-$K{})>=20'.format(number_rows+3,number_rows+3),
                                  'format': format_warn})
    
    #Plan exceeds capacity for over than 20%
    worksheet.conditional_format('G{}:G{}'.format(number_rows+4,number_rows+4),
                                {'type': 'cell',
                                'criteria': '>=',
                                'value': 1.2,
                                'format': format_alert})
    
    total_all1_fmt = workbook.add_format({**cmn.format_all, **{'align' : 'right', 'bold': True, 'top': 6}})
    total_effort1_fmt = workbook.add_format({**cmn.effort_fmt, **{'bold': True, 'top': 6}})
    total_percent1_fmt = workbook.add_format({**cmn.percent_fmt, **{'bold': True, 'top': 6}})
    total_all_fmt = workbook.add_format({**cmn.format_all, **{'align' : 'right', 'bold': True}})
    total_effort_fmt = workbook.add_format({**cmn.effort_fmt, **{'bold': True}})
    total_percent_fmt = workbook.add_format({**cmn.percent_fmt, **{'bold': True}})

    worksheet.write_string(number_rows+1, 0, sprintAggregatedReport['startTime'], total_all1_fmt)
    worksheet.write_string(number_rows+1, 1, sprintAggregatedReport['endTime'], total_all1_fmt)

    SprintText = sprintAggregatedReport['Sprint']

    if sprintAggregatedReport['state'] == 'CLOSED':
        SprintText = SprintText + ' (' + sprintAggregatedReport['state'] + ' ' + sprintAggregatedReport['completeTime'] + ')'
    else:
        SprintText = SprintText + ' (' + sprintAggregatedReport['state'] + ')'
        
    worksheet.write_string(number_rows+1, 2, SprintText, total_all1_fmt)
    worksheet.write_string(number_rows+1, 3, sprintAggregatedReport['Team'], total_all1_fmt)
    worksheet.write_string(number_rows+1, 5, 'Total:', total_all1_fmt)
    
    worksheet.write_string(number_rows+2, 5, 'Capacity/TERs:', total_all_fmt)
    worksheet.write_number(number_rows+2, 6, sprintAggregatedReport['Capacity'], total_effort_fmt)
    worksheet.write_number(number_rows+2, 7, sprintAggregatedReport['TERs'], total_effort_fmt) 
    worksheet.write_number(number_rows+2, 8, sprintAggregatedReport['RemainingCapacity'], total_effort_fmt)
    worksheet.write_string(number_rows+2, 9, '% done (qty):', total_all_fmt)
    worksheet.write_number(number_rows+2, 10, sprintAggregatedReport['CompletedTasksRate'], total_percent_fmt)
    
    worksheet.write_string(number_rows+3, 5, 'Plan, spent rates:', total_all_fmt)
    worksheet.write_number(number_rows+3, 6, sprintAggregatedReport['PlannedCapacityRate'], total_percent_fmt)    
    worksheet.write_number(number_rows+3, 7, sprintAggregatedReport['FactPlannedRate'], total_percent_fmt)
    worksheet.write_string(number_rows+3, 9, 'Curr % done (qty):', total_all_fmt)
    worksheet.write_number(number_rows+3, 10, sprintAggregatedReport['CurrCompletedTasksRate'], total_percent_fmt)

    worksheet.write_string(number_rows+4, 9, '% done (EV):', total_all_fmt)
    worksheet.write_number(number_rows+4, 10, sprintAggregatedReport['ImplementedRate'], total_percent_fmt)
    worksheet.write_string(number_rows+5, 9, 'Curr % done (EV):', total_all_fmt)
    worksheet.write_number(number_rows+5, 10, sprintAggregatedReport['CurrImplementedRate'], total_percent_fmt)

    worksheet.write_comment(xl_rowcol_to_cell(number_rows+1, 6),'Total planned estimates from Jira')
    worksheet.write_comment(xl_rowcol_to_cell(number_rows+1, 7),'Total spent time form Jira')  
    worksheet.write_comment(xl_rowcol_to_cell(number_rows+1, 8),'Total remaining estimates from Jira')   
    worksheet.write_comment(xl_rowcol_to_cell(number_rows+2, 6),'Team capacity for sprint according to resource plan')  
    worksheet.write_comment(xl_rowcol_to_cell(number_rows+2, 7),'Total TERs of team members for the period of sprint')  
    worksheet.write_comment(xl_rowcol_to_cell(number_rows+2, 8),'Remaining team capacity according to resource plan')  
    worksheet.write_comment(xl_rowcol_to_cell(number_rows+3, 6),'Planned estemates / Capacity rate')  
    worksheet.write_comment(xl_rowcol_to_cell(number_rows+3, 7),'Spent time / Original estimates ')   

    for column in [6, 7, 8]:
        worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, column),
                                "=SUM({:s}:{:s})".format(xl_rowcol_to_cell(1, column), xl_rowcol_to_cell(number_rows, column)),
                                total_effort1_fmt)
    
    for column in [9, 10, 14]:
        worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, column),
                                '=COUNTIF({:s}:{:s},"*")'.format(xl_rowcol_to_cell(1, column), xl_rowcol_to_cell(number_rows, column)),
                                total_all1_fmt)

    for column in [11, 12, 13]:
        worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, column),
                                '=COUNTIF({:s}:{:s},TRUE)'.format(xl_rowcol_to_cell(1, column), xl_rowcol_to_cell(number_rows, column)),
                                total_all1_fmt)
        
    worksheet.autofilter(0,0,number_rows+1,16)

    return

def format_TeamReport(workbook,
                      worksheet,
                      number_rows):

    format_all = workbook.add_format(cmn.format_all)
    header_fmt = workbook.add_format(cmn.header_fmt)
    effort_fmt = workbook.add_format(cmn.effort_fmt)
    percent_fmt = workbook.add_format(cmn.percent_fmt)

    worksheet.set_column('A:A', 20, format_all)
    worksheet.set_column('B:D', 7, format_all)
    worksheet.set_column('E:G', 9, percent_fmt)
    worksheet.set_column('H:J', 9, format_all)
    worksheet.set_column('K:L', 9, percent_fmt)
    worksheet.set_column('M:N', 9, format_all)
    worksheet.set_column('O:O', 9, percent_fmt)
    worksheet.set_column('P:R', 9, format_all)
    worksheet.set_column('S:T', 9, percent_fmt)
    worksheet.set_column('U:U', 9, format_all)

    worksheet.set_row(0, None, header_fmt)
    
    format_good = workbook.add_format(cmn.format_good)
    format_average = workbook.add_format(cmn.format_average)
    format_bad = workbook.add_format(cmn.format_bad)

    #KPI/THRESHOLD < 1
    worksheet.conditional_format('A2:A{}'.format(number_rows+1),
                                 {'type': 'formula', 'criteria': '=$U2<{}'.format(GREEN_THRESHOLD), 'format': format_good})
    worksheet.conditional_format('U2:U{}'.format(number_rows+1),
                                 {'type': 'formula', 'criteria': '=$U2<{}'.format(GREEN_THRESHOLD), 'format': format_good})

    worksheet.conditional_format('A2:A{}'.format(number_rows+1),
                                 {'type': 'formula', 'criteria': '=$U2<{}'.format(YELLOW_THRESHOLD), 'format': format_average})
    worksheet.conditional_format('U2:U{}'.format(number_rows+1),
                                 {'type': 'formula', 'criteria': '=$U2<{}'.format(YELLOW_THRESHOLD), 'format': format_average})

    
    worksheet.conditional_format('E2:E{}'.format(number_rows+1),
                                 {'type': 'formula', 'criteria': '=OR($E2<0.5,$E2>1.4)', 'format': format_bad})
    worksheet.conditional_format('F2:F{}'.format(number_rows+1),
                                 {'type': 'formula', 'criteria': '=OR($F2<0.7,$F2>1.3)', 'format': format_bad})

    worksheet.conditional_format('E2:E{}'.format(number_rows+1),
                                 {'type': 'formula', 'criteria': '=OR($E2<0.6,$E2>1.3)', 'format': format_average})
    worksheet.conditional_format('F2:F{}'.format(number_rows+1),
                                 {'type': 'formula', 'criteria': '=OR($F2<0.9,$F2>1.2)', 'format': format_average})

    worksheet.conditional_format('E2:E{}'.format(number_rows+1),
                                 {'type': 'formula', 'criteria': '=AND($E2>=0.6,$E2<=1.3)', 'format': format_good})
    worksheet.conditional_format('F2:F{}'.format(number_rows+1),
                                 {'type': 'formula', 'criteria': '=AND($F2>=0.9,$F2<=1.2)', 'format': format_good})


    worksheet.conditional_format('G2:G{}'.format(number_rows+1), cmn.format_bar_green)
    worksheet.conditional_format('O2:O{}'.format(number_rows+1), cmn.format_bar_green)
    worksheet.conditional_format('S2:S{}'.format(number_rows+1), cmn.format_bar_green)
    worksheet.conditional_format('T2:T{}'.format(number_rows+1), cmn.format_bar_green)
    worksheet.conditional_format('K2:K{}'.format(number_rows+1), cmn.format_bar_blue)
    worksheet.conditional_format('L2:L{}'.format(number_rows+1), cmn.format_bar_blue)


    worksheet.write_comment(xl_rowcol_to_cell(0, 1),'Capacity for sprint from Resource Plan')
    worksheet.write_comment(xl_rowcol_to_cell(0, 2),'Sum of Remaining Estimates at sprint start')    
    worksheet.write_comment(xl_rowcol_to_cell(0, 3),'Sum of Time Spent for sprint tasks')  
    worksheet.write_comment(xl_rowcol_to_cell(0, 4),'Plan / Capacity rate')
    worksheet.write_comment(xl_rowcol_to_cell(0, 5),'Fact / Plan rate')     
    worksheet.write_comment(xl_rowcol_to_cell(0, 6),'[Sum of resolved tasks Estimations] / Plan')  
    worksheet.write_comment(xl_rowcol_to_cell(0, 7),'Remaining capacity for sprint from Resource Plan')
    worksheet.write_comment(xl_rowcol_to_cell(0, 8),'Sum of TERs for project GP code from ERP for sprint period')  
    worksheet.write_comment(xl_rowcol_to_cell(0, 9),'Fact / TERs rate')
    worksheet.write_comment(xl_rowcol_to_cell(0, 10),'[Sum of Fact of resolved tasks] / TERs rate')   
    worksheet.write_comment(xl_rowcol_to_cell(0, 11),'Total number of assigned tasks in sprint')   
    worksheet.write_comment(xl_rowcol_to_cell(0, 12),'Total number of resolved tasks in sprint')   
    worksheet.write_comment(xl_rowcol_to_cell(0, 13),'% of resolved tasks')    
    worksheet.write_comment(xl_rowcol_to_cell(0, 14),'Total number of initially planned tasks in sprint')   
    worksheet.write_comment(xl_rowcol_to_cell(0, 15),'Total number of tasks added after sprint start')   
    worksheet.write_comment(xl_rowcol_to_cell(0, 16),'Total number of descoped tasks')   
    worksheet.write_comment(xl_rowcol_to_cell(0, 17),'Current % of resolved tasks (as of now, may be resolved after sprint end)')   
    worksheet.write_comment(xl_rowcol_to_cell(0, 18),'Current [Sum of resolved tasks Estimations] / Plan (as of now, may be resolved after sprint end)')
    worksheet.write_comment(xl_rowcol_to_cell(0, 19),'KPI based on P/C, F/P and Done EV')   

    return

def writeReport(path,
                sprintDetailedReport,
                sprintAggregatedReport,
                sprintTeamReport,
                Detailed_columns,
                Detailed_labels,
                Aggregated_columns,
                Aggregated_labels,
                Team_columns,
                Team_labels):

    writer = pd.ExcelWriter(path, engine='xlsxwriter')

    pd.io.formats.excel.header_style = None

    sprintDetailedReport.to_excel(writer,
                                  index = False,
                                  freeze_panes=(1,2),
                                  sheet_name = 'Details',
                                  columns = Detailed_columns,
                                  header = Detailed_labels)

    workbook = writer.book
    
    format_DetailedReport(workbook,
                          writer.sheets['Details'],
                          len(sprintDetailedReport.index),
                          sprintAggregatedReport,
                          Aggregated_columns,
                          Aggregated_labels)
    
    sprintTeamReport.to_excel(writer,
                              index_label = 'Name',
                              index = True,
                              freeze_panes = (0,1),
                              sheet_name = 'Team',
                              columns = Team_columns,
                              header = Team_labels)
    
    format_TeamReport(workbook,
                      writer.sheets['Team'],
                      len(sprintTeamReport.index))

    writer.save()

    return

def sprintReport(project,
                 filename,
                 sprint,
                 jira,
                 path = cmn.DEFAULT_PATH + 'Documents\\',
                 WBS_path = '',
                 TER_path = '',
                 getdailyrep = False):

    print('Preparing ' + filename)
    
    board_id = sprint['Board']

    Burndown = getScopeChangeBurndownChart(jira, board_id, sprint.name)
    
    print('Loaded Burndown')

    parsedBurndown = parseBurndown(Burndown)

    sprints_dates = getSprintsDates(project)

    sprintDetailedReport = detailedSprintReport(jira, parsedBurndown, Burndown, getdailyrep)

    if sprintDetailedReport is None:
    
        print('No issues in sprint ' + sprint['Sprint_name'] + ' on board ' + str(board_id))

        return
    
    print('sprintDetailedReport ready')

    sprintTeamReport = teamSprintReport(project, sprint, sprints_dates, jira, sprintDetailedReport)
    
    print('sprintTeamReport ready')

    sprintAggregatedReport = aggregatedSprintReport(project, sprint, sprintDetailedReport, Burndown, sprints_dates)

    print('sprintAggregatedReport ready')
    
    Detailed_columns = ['parent',
                        'key',
                        'Summary',
                        'Assignee',
                        'Status',
                        'Status_mapped',
                        'initialEstimate',
                        'timeSpent',
                        'remainingEstimate',
                        'added_str',
                        'completed_str',
                        'initialScope',
                        'additionalScope',
                        'descope',
                        'descoped',
                        'notParent',
                        'dailyreport']
    Detailed_labels = ['parent',
                        'key',
                        'Summary',
                        'Assignee',
                        'Status',
                        'Status2',
                        'Estimate',
                        'Spent',
                        'Remaining',
                        'added',
                        'completed',
                        'initScope',
                        'addScope',
                        'descope',
                        'descoped',
                        'notParent',
                        'Daily Report']

    Aggregated_columns = ['Team',
                        'Sprint',
                        'startTime',
                        'endTime',
                        'state',
                        'completeTime',
                        'Capacity',
                        'RemainingCapacity',
                        'PlannedTotal',
                        'Fact',
                        'TERs',
                        'ImplementedRate',
                        'FactPlannedRate',
                        'PlannedCapacityRate',
                        'InitiallyPlannedTasks',
                        'AdditionalScope',
                        'DescopedTasks',
                        'PlannedTasksTotal',
                        'CompletedTasks',
                        'CompletedTasksRate']
    Aggregated_labels = ['Team',
                        'Sprint',
                        'startTime',
                        'endTime',
                        'state',
                        'completeTime',
                        'Capacity',
                        'RemainingCapacity',
                        'PlannedTotal',
                        'Fact',
                        'TERs',
                        'ImplementedRate',
                        'FactPlannedRate',
                        'PlannedCapacityRate',
                        'InitiallyPlannedTasks',
                        'AdditionalScope',
                        'DescopedTasks',
                        'PlannedTasksTotal',
                        'CompletedTasks',
                        'CompletedTasksRate']

    Team_columns = ['Capacity',
                    'PlannedTotal',
                    'Fact',
                    'PlannedCapacityRate',
                    'FactPlannedRate',
                    'ImplementedRate',
                    'RemainingCapacity',
                    'RemainingTotal',
                    'TERs',
                    'Velocity',
                    'Efficiency',
                    'PlannedTasksTotal',
                    'CompletedTasks',
                    'CompletedTasksRate',
                    'InitiallyPlannedTasks',
                    'AdditionalScope',
                    'DescopedTasks',
                    'CurrCompletedTasksRate',
                    'CurrImplementedRate',
                    'KPI']

    Team_labels = ['Capacity',
                   'Plan',
                   'Fact',
                   'Plan / Capacity',
                   'Fact / Plan',
                   'Done EV, %',
                   'Remaining Capacity',
                   'Remaining',
                   'TERs',
                   'Velocity',
                   'Efficiency',
                   'Plan QTY',
                   'Done QTY',
                   'Done, %',
                   'Init Plan',
                   'Additional',
                   'Descoped',
                   'Done Current, %',
                   'Done Curr EV, %',
                   'KPI']

    writeReport(path + filename,
                sprintDetailedReport,
                sprintAggregatedReport,
                sprintTeamReport,
                Detailed_columns,
                Detailed_labels,
                Aggregated_columns,
                Aggregated_labels,
                Team_columns,
                Team_labels)
    
    print('Done!')

    return

