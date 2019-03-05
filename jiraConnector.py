from jira import JIRA
from jira.client import GreenHopper
import pandas as pd
import common as cmn
import re

statuses_mapped = pd.read_excel(cmn.DEFAULT_PATH + 'Documents\\statuses.xlsx',index_col=0)

def get_pwd():
    pwd = ''
    return pwd

def psup_login():
    psup = JIRA(server='https://psup.netcracker.com', basic_auth =('rudu0916', get_pwd()), max_retries=0)

    fields = psup.fields()
    setattr(psup, 'sprintfield', [x['id'] for x in fields if x['name'] == 'Sprint'][0])
    setattr(psup, 'epiclinkfield', [x['id'] for x in fields if x['name'] == 'Epic Link'][0])
    setattr(psup, 'featurelinkfield', [x['id'] for x in fields if x['name'] == 'Feature Link'][0])

    return psup

def tms_login():
    tms = JIRA(server='https://tms.netcracker.com', basic_auth =('rudu0916', get_pwd()))

    fields = tms.fields()
    setattr(tms, 'sprintfield', [x['id'] for x in fields if x['name'] == 'Sprint'][0])
    setattr(tms, 'epiclinkfield', [x['id'] for x in fields if x['name'] == 'Epic Link'][0])

    return tms

def parse_sprints(s):
    return s[s.find(',name=')+6:s.find(',',s.find(',name=')+1)]

def get_story_parent(raw_story, jira):
    parent = getattr(raw_story.fields,jira.featurelinkfield,None)
    if parent is None:
        return getattr(raw_story.fields,jira.epiclinkfield)
    else:
        return parent

def get_assignee(assignee):
    if assignee is None:
        return 'Unassigned'
    else:
        return assignee.displayName

def get_status_mapped(issue):
    status = issue.fields.status.name
    if (status == 'Closed') and (issue.fields.resolution is not None):
        if issue.fields.resolution.name in ['Canceled by Requester','Cannot Reproduce']:
            return 'Cancelled'
    return statuses_mapped.loc[status]['Type']

def get_plan(issue, WBSitem = 'HL Estimate, md', WBStag = 'WBS'):
    
    if cmn.nvl(WBStag,'') == '':
        return None

    dsc = issue.fields.description
    
    if dsc is None:
        return None
    
    m = re.search('#' + WBStag + '(\{.*\})',dsc)
    
    if m is None:
        return None

    WBSdata = eval(m.group(1).replace('nan','0.0'))

    #'Dev LOE' is legacy, remove later
    if WBSitem in WBSdata:
        return WBSdata[WBSitem] * 8
    elif 'Dev LOE' in WBSdata:
        return WBSdata['Dev LOE'] * 8
    else:
        return None

def get_step(issue):
    if re.search('^\W*QA.+',issue.fields.summary) is None:
        return 'Dev'
    else:
        return 'QA'

def get_priority(issue):
    if issue.fields.priority is None:
        return ''
    else:
        return issue.fields.priority.name

def getFRTeams(issue):
    if 'customfield_23024' in issue.raw['fields']:
        teams = issue.raw['fields']['customfield_23024']
        if teams is not None:
            return list(map(lambda x: x['value'], issue.raw['fields']['customfield_23024']))

    return []

def parse_components(components):
    c = set()
    for component in components:
        c.add(component.name)
    return c

def getTeam(issue):
    team = ''
    if 'customfield_23025' in issue.raw['fields']:
        if 'value' in issue.raw['fields']['customfield_23025']:
            team = issue.raw['fields']['customfield_23025']['value']

    if team == '':
        components = parse_components(issue.fields.components)
        team = ', '.join(components)

    return team

def getCustomers(issue):
    try:
        customers = ', '.join([str(customer.value) for customer in cmn.nvl(issue.fields.customfield_21820,[])])
    except AttributeError:
        customers = ''
    return customers

def get_FRs(jira, filter_id, WBStag = 'WBS'):

    raw_FRs = jira.search_issues(jira.filter(filter_id).jql,maxResults=1000)

    FRs_keys = list([x.key for x in raw_FRs])

    FRs_list = list(map(lambda x:
                    [cmn.nvl(get_plan(x, WBStag=WBStag, WBSitem='HL Estimate, md'), cmn.nvl(x.fields.customfield_23121,0)),
                    ', '.join([parse_sprints(sprint) for sprint in cmn.nvl(getattr(x.fields,jira.sprintfield),[])]),
                    cmn.nvl(get_plan(x, WBStag=WBStag, WBSitem='Analysis LOE, md'), cmn.nvl(x.fields.customfield_23422,0)),
                    cmn.nvl(get_plan(x, WBStag=WBStag, WBSitem='Build LOE, md'), cmn.nvl(x.fields.customfield_23424,0)),
                    cmn.nvl(get_plan(x, WBStag=WBStag, WBSitem='Design LOE, md'), cmn.nvl(x.fields.customfield_23423,0)),
                    cmn.nvl(get_plan(x, WBStag=WBStag, WBSitem='QA LOE, md'), cmn.nvl(x.fields.customfield_23426,0)),
                    cmn.nvl(get_plan(x, WBStag=WBStag, WBSitem='Bug Fixing LOE, md'), cmn.nvl(x.fields.customfield_23425,0)),
                    x.fields.summary,
                    getCustomers(x),
                    x.fields.duedate,
                    ', '.join([str(label) for label in cmn.nvl(x.fields.labels,[])]),
                    ', '.join([str(fixVersion.name) for fixVersion in cmn.nvl(x.fields.fixVersions,[])]),
                    x.fields.status.name,
                    get_status_mapped(x),
                    ', '.join(getFRTeams(x))],
                    raw_FRs))

    FRs_df = pd.DataFrame(FRs_list,
                          index=FRs_keys,
                          columns=['HL Estimate, md',
                                   'Sprint',
                                   'Analysis LOE, md',
                                   'Build LOE, md',
                                   'Design LOE, md',
                                   'QA LOE, md',
                                   'Bug Fixing LOE, md',
                                   'Summary',
                                   'Customers',
                                   'Due Date',
                                   'Labels',
                                   'Fix Version/s',
                                   'Status',
                                   'Status Mapped',
                                   'Teams'])

    return FRs_df

def get_Stories(jira, filter_id='', FRs=[], add_clause='', WBStag = 'WBS'):

    if len(filter_id) > 0:
        jql = jira.filter(filter_id).jql
    else:
        jql = '"Feature Link" in (' + ','.join(str(key) for key in FRs) + ')'
        if len(add_clause) > 0:
            jql = jql + ' AND ' + add_clause

    raw_Stories = jira.search_issues(jql,maxResults=1000)
    
    Stories_keys = [[get_story_parent(x, jira) for x in raw_Stories],
                    [x.key for x in raw_Stories]]

    Stories_list = list(map(lambda x:
                    [', '.join([parse_sprints(sprint) for sprint in cmn.nvl(getattr(x.fields,jira.sprintfield),[])]),
                    x.fields.summary,
                    x.fields.duedate,
                    ', '.join([str(label) for label in cmn.nvl(x.fields.labels,[])]),
                    ', '.join([str(fixVersion.name) for fixVersion in cmn.nvl(x.fields.fixVersions,[])]),
                    cmn.nvl(x.fields.aggregatetimeestimate,0)/3600,
                    cmn.nvl(x.fields.timeoriginalestimate,0)/3600,
                    cmn.nvl(x.fields.timeestimate,0)/3600,
                    cmn.nvl(x.fields.aggregatetimeoriginalestimate,0)/3600,
                    get_plan(x, WBStag=WBStag),
                    cmn.nvl(x.fields.aggregatetimespent,0)/3600,
                    cmn.nvl(x.fields.timespent,0)/3600,
                    ', '.join([str(subtask) for subtask in cmn.nvl(x.fields.subtasks,[])]),
                    x.fields.status.name,
                    get_status_mapped(x),
                    get_assignee(x.fields.assignee),
                    getTeam(x)],
                    raw_Stories))

    Stories_df = pd.DataFrame(Stories_list,
                              index=Stories_keys,
                              columns=[ 'Sprint',
                                        'Summary',
                                        'Due Date',
                                        'Labels',
                                        'Fix Version/s',
                                        'Σ Remaining Estimate',
                                        'Original Estimate',
                                        'Remaining Estimate',
                                        'Σ Original Estimate',
                                        'WBS Dev LOE',
                                        'Σ Time Spent',
                                        'Time Spent',
                                        'Sub-Tasks',
                                        'Status',
                                        'Status Mapped',
                                        'Assignee',
                                        'Team'])

    return Stories_df

def get_WorkItems(jira, jql=''):

    raw_WorkItems = jira.search_issues(jql,maxResults=1000)
    
    WorkItems_keys = [[getattr(x.fields,jira.epiclinkfield) for x in raw_WorkItems],
                    [x.key for x in raw_WorkItems]]

    WorkItems_list = list(map(lambda x:
                    [', '.join([parse_sprints(sprint) for sprint in cmn.nvl(getattr(x.fields,jira.sprintfield),[])]),
                    ', '.join([str(component.name) for component in cmn.nvl(x.fields.components,[])]),
                    x.fields.summary,
                    x.fields.duedate,
                    ', '.join([str(label) for label in cmn.nvl(x.fields.labels,[])]),
                    ', '.join([str(fixVersion.name) for fixVersion in cmn.nvl(x.fields.fixVersions,[])]),
                    cmn.nvl(x.fields.aggregatetimeestimate,0)/3600,
                    cmn.nvl(x.fields.timeoriginalestimate,0)/3600,
                    cmn.nvl(x.fields.timeestimate,0)/3600,
                    cmn.nvl(x.fields.aggregatetimeoriginalestimate,0)/3600,
                    cmn.nvl(x.fields.aggregatetimespent,0)/3600,
                    cmn.nvl(x.fields.timespent,0)/3600,
                    ', '.join([str(subtask) for subtask in cmn.nvl(x.fields.subtasks,[])]),
                    x.fields.status.name,
                    x.fields.priority.name,
                    get_status_mapped(x),
                    get_assignee(x.fields.assignee)],
                    raw_WorkItems))

    WorkItems_df = pd.DataFrame(WorkItems_list,
                              index=WorkItems_keys,
                              columns=[ 'Sprint',
                                        'Components',
                                        'Summary',
                                        'Due Date',
                                        'Labels',
                                        'Fix Version/s',
                                        'Σ Remaining Estimate',
                                        'Original Estimate',
                                        'Remaining Estimate',
                                        'Σ Original Estimate',
                                        'Σ Time Spent',
                                        'Time Spent',
                                        'Sub-Tasks',
                                        'Status',
                                        'Priority',
                                        'Status Mapped',
                                        'Assignee'])

    return WorkItems_df

def get_subtasks(jira, parents):

    jql = 'parent in (' + ','.join(str(key) for key in parents.index.levels[1]) + ')'
    
    raw_subtasks = jira.search_issues(jql,maxResults=1000)

    subtasks_keys = [[parents[parents.index.get_level_values(1) == x.fields.parent.key].iloc[0].name[0] for x in raw_subtasks],
                     [x.fields.parent.key for x in raw_subtasks],
                    [x.key for x in raw_subtasks]]

    subtasks_list = list(map(lambda x:
                    [', '.join([parse_sprints(sprint) for sprint in cmn.nvl(getattr(x.fields,jira.sprintfield),[])]),
                    ', '.join([str(component.name) for component in cmn.nvl(x.fields.components,[])]),
                    x.fields.summary,
                    x.fields.duedate,
                    ', '.join([str(label) for label in cmn.nvl(x.fields.labels,[])]),
                    get_step(x),
                    ', '.join([str(fixVersion.name) for fixVersion in cmn.nvl(x.fields.fixVersions,[])]),
                    cmn.nvl(x.fields.timeoriginalestimate,0)/3600,
                    cmn.nvl(x.fields.timeestimate,0)/3600,
                    cmn.nvl(x.fields.timespent,0)/3600,
                    x.fields.status.name,
                    get_priority(x),
                    get_status_mapped(x),
                    get_assignee(x.fields.assignee)],
                    raw_subtasks))

    subtasks_df = pd.DataFrame(subtasks_list,
                               index = subtasks_keys,
                              columns=['Sprint',
                                        'Components',
                                        'Summary',
                                        'Due Date',
                                        'Labels',
                                        'Step',
                                        'Fix Version/s',
                                        'Original Estimate',
                                        'Remaining Estimate',
                                        'Time Spent',
                                        'Status',
                                        'Priority',
                                        'Status Mapped',
                                        'Assignee'])

    return subtasks_df

def update_statuses(jira, target):
    raw_statuses = jira.statuses()

    statuses_mapped = pd.read_excel(cmn.DEFAULT_PATH + 'Documents\\statuses.xlsx',index_col=0)

    statuses_names = list([x.name for x in raw_statuses])

    statuses_list = list(map(lambda x:
                    dict(x.raw).get('statusCategory').get('name'),
                    raw_statuses))

    statuses_df = pd.DataFrame(statuses_list,
                            index=statuses_names,
                            columns=['statusCategory'])

    def status_m(name):
        try:
            return statuses_mapped.loc[name]['Type']
        except KeyError:
            return 'N/A'

    statuses_df1 = statuses_df.join(pd.DataFrame(statuses_df.apply(lambda x: status_m(x.name), axis=1)))

    statuses_df1.to_excel(cmn.DEFAULT_PATH + 'Documents\\' + target)

    return