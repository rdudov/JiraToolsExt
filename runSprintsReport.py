import sprintsReport as sp
import jiraConnector as jc
import datetime as dt
import re

def boardsByProject(project):
    if project in ['VDOCSMR6','VDOCSMR7']:
        boards = [901]
    elif project in ['VDOCDPR6', 'VDOCDPR7']:
        boards = [911]
    elif project == 'CM92':
        boards = [777]
    elif project in ['ES20_R4']:
        #<JIRA Board: name='Ecosystem 2.0 Development', id=824>
        boards = [834]
    elif project in ['ESO91']:
        #<JIRA Board: name='[FT: BETA] HOM R9.0 Beta', id=736>
        boards = [736]
        #<JIRA Board: name='MANO ESO Codebase (Tau)', id=783>
        boards = boards + [783]
        #<JIRA Board: name='MANO ESO Codebase (Phi)', id=800>
        boards = boards + [800]
    else: #OLD
        if project == 'VDOC_R3':
            #<JIRA Board: name='VDOC.R3 Service Modeling', id=604>
            boards = [604]
        elif project in ['CM91','CM911']:
            boards = [777]
        elif project == 'VDOCR5':
            boards = [826]
        elif project in ['ES20_R3']:
            #<JIRA Board: name='Ecosystem 2.0 Development', id=724>
            boards = [724]
        elif project == 'VDOCR4':
            boards = [707,757]
        elif project in ['UUI_81','UUI_90']:
            #<JIRA Board: name='UMBRIU Board', id=525>
            boards = [525]
            #<JIRA Board: name='MANO Graphical Views', id=510>
            boards = boards + [510]
            #<JIRA Board: name='[FT: Omega] HOM R9.0 Beta', id=715>
            boards = boards + [510]
        elif project in ['IPAM_81','IPAM_90']:
            #<JIRA Board: name='IPAM Board', id=507>
            boards = [507]
        elif project in ['TFA_81','TFA_90']:
            #<JIRA Board: name='HOMTFA Board', id=511>
            boards = [511]
        elif project in ['ES20_R1','ES20_R2']:
            #<JIRA Board: name='Ecosystem 2.0 Development', id=724>
            boards = [724]
    return boards

def printSprints(project,jira):
    for board_id in boardsByProject(project):
        sprints = sp.getSprints(jira, board_id, project)
        print(sprints.to_string())

def runReportForSprint(project, board_id, sprint_id, jira = None, getdailyrep = False):
    if jira is None:
        jira = jc.psup_login()
        print('Login OK')

    sprints = sp.getSprints(jira, board_id, project)
    sprint = sprints.loc[sprint_id]
    filename = 'sprintReport_{}_{}_{}_{}.xlsx'.format(project,sprint['Team'],sprint['Sprint'],dt.datetime.now().strftime("%Y-%m-%d"))
    filename = re.sub('[\/:*?"<>|]+','_',filename)
    sp.sprintReport(project, filename, sprint, jira, getdailyrep=getdailyrep)
    return


def runReport(project, jira = None, sprint = 0, getdailyrep = False):

    if jira is None:
        jira = jc.psup_login()
        print('Login OK')

    for board_id in boardsByProject(project):
        sprints = sp.getSprints(jira, board_id, project)
        sprints_filtered = sp.getSprintIDs(sprints, sprint)

        if len(sprints_filtered) == 0:
            print('No active sprints for board {}!'.format(board_id))
            continue

        for sprint_id in sprints_filtered.index.get_values():
            runReportForSprint(project, board_id, sprint_id, jira, getdailyrep=getdailyrep)

    return

jira = jc.psup_login()

#runReportForSprint('CM91',777,989,jira)

#printSprints('VDOCDPR6',jira)
#runReportForSprint('VDOCR6',901,1223,jira)
#runReport('VDOCDPR6',jira)

#printSprints('VDOCSMR7',jira)
#runReportForSprint('VDOCR7',901,1337,jira)
runReport('VDOCSMR7',jira)

#printSprints('ES20_R4',jira)
#runReport('ES20_R4',jira, getdailyrep=True)
#runReportForSprint('ES20_R4',834,1125,jira, getdailyrep=True)
#runReportForSprint('ES20_R4',834,1152,jira, getdailyrep=False)

#printSprints('CM92',jira)
#runReport('CM92',jira)
#runReportForSprint('CM92',777,1249,jira)

#printSprints('ESO91',jira)
#runReport('ESO91',jira)
#runReportForSprint('ESO91',736,1066,jira) #Beta

#runReportForSprint('ESO91',783,1155,jira) #Tau
#runReportForSprint('ESO91',800,1155,jira) #Phi
