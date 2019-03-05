import pandas as pd
import jiraConnector as jc
import common as cmn
import datetime as dt
import pytz

jira = jc.psup_login()

jql = '"Epic link" in (CHOM-2423,CHOM-2344)'

def parse_components(components):
    c = set()
    for component in components:
        c.add(component.name)
    return c

def alignComponents(jql, jira, update, add_clause=""):

    stories = jira.search_issues(jql,maxResults=1000)

    print('Loaded {} stories'.format(len(stories)))

    stories_keys = list([x.key for x in stories])
    stories_list = list(map(lambda x:
                        [parse_components(x.fields.components)],
                        stories))
    stories_df = pd.DataFrame(stories_list,
                                index=stories_keys,
                                columns=['components'])

    subtasks_jql = 'parent in (' + ','.join(str(key) for key in stories_keys) + ') and created > -10d'
    
    if add_clause != "":
        subtasks_jql = subtasks_jql + ' AND ' + add_clause

    subtasks = jira.search_issues(subtasks_jql,maxResults=1000)
    
    print('Loaded {} subtasks'.format(len(subtasks)))

    subtasks_keys = list([x.key for x in subtasks])
    subtasks_list = list(map(lambda x:
                        [x.fields.parent.key,
                         ', '.join([jc.parse_sprints(sprint) for sprint in cmn.nvl(getattr(x.fields,jira.sprintfield),[])]),
                         x.fields.status.name,
                         parse_components(x.fields.components)],
                        subtasks))
    subtasks_df = pd.DataFrame(subtasks_list,
                                index=subtasks_keys,
                                columns=['parent','sprints','status','components'])

    subtasks_df = subtasks_df.join(stories_df,on='parent',rsuffix='_story')

    subtasks_df['diff'] = subtasks_df.apply(lambda row: row.components^row.components_story,axis=1)
    subtasks_df['cmn'] = subtasks_df.apply(lambda row: row.components&row.components_story,axis=1)
    
    print('KEY       STATUS   SUBTASK    STORY    DIFF    CMN')

    for (key,subtask) in subtasks_df.iterrows():
        if subtask.sprints != "" and subtask.status != 'Closed':
            if len(subtask['diff']) > 0 and len(subtask['cmn']) == 0:
                print('{} {} {} {} {} {}'.format(key,
                                            str(subtask.status),
                                            str(subtask.components),
                                            str(subtask.components_story),
                                            str(subtask['diff']),
                                            str(subtask['cmn'])))
                if update and len(subtask.components) == 0:
                    if len(subtask['diff']) > 1:
                        print('Check {} multiple component {}'.format(key,str(subtask['diff'])))
                    else:
                        issue = jira.issue(key)
                        issue.update(notify=False,
                                     fields={"components": [{'name': component} for component in subtask['diff']]})
                        print('Updated {} set component {}'.format(key,str(subtask['diff'])))



alignComponents(jql, jira, False)
alignComponents(jql, jira, True)


def resetRemaining(jira, update):

    jql = '"Epic link" in (CHOM-2423,CHOM-2344)'
    
    stories = jira.search_issues(jql,maxResults=1000)

    print('Loaded {} stories'.format(len(stories)))

    stories_keys = list([x.key for x in stories])

    subtasks_jql = 'parent in (' + ','.join(str(key) for key in stories_keys) + ') and resolved > -10d and remainingEstimate > 0 and status = Resolved'
    
    subtasks = jira.search_issues(subtasks_jql,maxResults=1000)

    subtasks_keys = list([x.key for x in subtasks])
    subtasks_list = list(map(lambda x:
                        [x.fields.timeestimate/3600],
                        subtasks))
    subtasks_df = pd.DataFrame(subtasks_list,
                                index=subtasks_keys,
                                columns=['remainingEstimate'])

    if len(subtasks_df) == 0:
        print('Nothing to update')

    for (key,subtask) in subtasks_df.iterrows():
        print('{} {}'.format(key, str(remainingEstimate)))

        if update:
            issue = jira.issue(key)
            issue.update(notify=False, remainingEstimate=0)
            print('Updated {}'.format(key))

            
resetRemaining(jira, False)
resetRemaining(jira, True)




project = 'VDOCR5'
FR_filter = '59394'
add_clause = 'fixVersion = "Release 5"'
filename = 'report_' + project + '_' + datetime.datetime.now().strftime("%Y-%m-%d") + '.xlsx'
Stories_filter = ''
WBSpath = ''
timesheet = ter.getTimesheetForProject(project)
WBStag = project + '_WBS'
jira = jc.psup_login()
FRs = jc.get_FRs(jira, FR_filter, WBStag=WBStag)
#Stories = jc.get_Stories(jira, FRs=FRs.index, add_clause=add_clause, WBStag=WBStag)
FRs=FRs.index
jql = '"Feature Link" in (' + ','.join(str(key) for key in FRs) + ')'
jql = jql + ' AND ' + add_clause
raw_Stories = jira.search_issues(jql,maxResults=1000)


Stories_c = Stories[Stories["Sub-Tasks"].str.len() > 0]
Stories.apply(lambda row: calc_subtasks(row.name[1],subtasks),axis=1)
Stories[Stories["Sub-Tasks"].str.len() > 0].apply(lambda row: calc_subtasks(row.name[1],subtasks),axis=1)

Stories.loc[Stories.index.get_level_values(1) == 'HOMTFA-1849']["Sub-Tasks"]
Stories.loc[Stories.index.get_level_values(1) == 'HOMTFA-1849'].apply(lambda row: calc_subtasks(row.name[1],subtasks),axis=1)

Stories[Stories.index.get_level_values(1) == 'HOMTFA-1849'].iloc[0].name[0]

subtasks[subtasks.index.get_level_values(0).isin(['HOMTFA-1849'])]
subtasks.filter(like='HOMTFA-2018',axis=0)

sum(subtasks_c[(subtasks_c['Step']=='QA') & (subtasks_c['Status Mapped']=='Done')].apply(original_estimate,axis=1))

filteredBurndownNotDone = parsedBurndown.query('timestamp >= ' + str(lastDone) +
                                                ' and timestamp <= ' + str(endTime) +
                                                ' and notDone == True '
                                                ' and key == "' + key + '"')

m = re.search('(\[.+\])','[Func][DEV] Dev Lead')
m = re.search('([A-Za-z]+?)(\[.+\])','[Func][DEV] Dev Lead')

m = re.search('(ab+.+cd)+.','[ab][cd] abcd')
m = re.search('(\[\W+\]).+(\[.+?\])','[ab][cd][*;] abcd')
m = re.search('(\[\w+\])?','[ab1][cd][*;] abcd')

m = re.search('\[\w+\].*(\[\w+\])','[ab1][cd][*;] abcd')

m = re.search('(\w+)','[ab1][cd][*;] abcd')


with open('C:\\Users\\rudu0916\\Documents\\usr', 'wb') as f:
        f.write(b'083')
        f.write(b'117')
        f.write(b'109')
        f.write(b'109')
        f.write(b'101')
        f.write(b'114')
        f.write(b'036')
        f.write(b'104')
        f.write(b'105')
        f.write(b'110')
        f.write(b'105')
        f.write(b'110')
        f.write(b'103')


a = ''
for c in a:
    ord(c)

raw_statuses = jira.statuses()

statuses_names = list([x.name for x in raw_statuses])

statuses_list = list(map(lambda x:
                dict(x.raw).get('statusCategory').get('name'),
                raw_statuses))

statuses_df = pd.DataFrame(statuses_list,
                        index=statuses_names,
                        columns=['statusCategory'])

    return FRs_df

def status_m(name):
    try:
        return statuses_mapped.loc[name]['Type']
    except KeyError:
        return 'N/A'

statuses_df1 = statuses_df.join(pd.DataFrame(statuses_df.apply(lambda x: status_m(x.name), axis=1)))

statuses_df1.to_excel('C:\\Users\\rudu0916\\Documents\\jira_statuses.xlsx')


r = requests.get(runReportRequest, auth=('rudu0916', jc.get_pwd()), verify=False)

print(r.status_code)
print(r.headers['content-type'])


from urllib2 import Request, build_opener, HTTPCookieProcessor, HTTPHandler
import httplib, urllib, cookielib, Cookie, os

conn = httplib.HTTPConnection('webapp.pucrs.br')

#COOKIE FINDER
cj = cookielib.CookieJar()
opener = build_opener(HTTPCookieProcessor(cj),HTTPHandler())
req = Request('http://webapp.pucrs.br/consulta/principal.jsp')
f = opener.open(req)
html = f.read()
for cookie in cj:
    c = cookie
#FIM COOKIE FINDER

params = urllib.urlencode ({'pr1':111049631, 'pr2':<pass>})
headers = {"Content-type":"text/html",
           "Set-Cookie" : "JSESSIONID=70E78D6970373C07A81302C7CF800349"}
            # I couldn't set the value automaticaly here, the cookie object can't be converted to string, so I change this value on every session to the new cookie's value. Any solutions?

conn.request ("POST", "/consulta/servlet/consulta.aluno.ValidaAluno",params, headers) # Validation page
resp = conn.getresponse()

temp = conn.request("GET","/consulta/servlet/consulta.aluno.Publicacoes") # desired content page
resp = conn.getresponse()

print resp.read()

cookie = 'isInProcess=true; SERVERID=node1; ZP_CAL=%27fdow%27%3Anull%2C%27history%27%3A%222017/06/22/12/49%22%2C%27sortOrder%27%3A%22asc%22%2C%27hsize%27%3A9; netcracker_from_url_9f153058-bcbc-4b96-85a9-81f63e9dc79f=/ncobject.jsp?id=9147386565313114588; netcracker_from_url_efbdc056-7ecc-4b9e-9984-2cc099203082=/report/exportreport.jsp?param0=9147133456513389207&expobject=7100462886013127576&param1=%2712%2F18%2F2016%27&param2=%2706%2F22%2F2017%27&param3=NULL&param4=NULL&adapterId=9135923224913544218&param5=%27%25%27&report=7100462886013127576&param6=NULL&parent=7040459919013415288&object=3112170189013929210; JSESSIONID=jN0OCMYVoDhfPHKzX02lVaMuZzpq-iy5hI5sHmExRMHNHqy1erHP!-2107866838'