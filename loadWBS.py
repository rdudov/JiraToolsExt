import jiraConnector as jc
import common as cmn
import pandas as pd
import WBStools as wbs
import re

jira = jc.psup_login()

#UUI_FRs = jc.get_FRs(jira, '55044')

VDOCR5_WBSpath = 'C:\\Users\\rudu0916\\Desktop\\SDN_NFV\\AVP\\SM\\R5\\R5 proposal\VDOC_R5_WBS v1.xlsx'
du_vCPE_WBSpath = 'C:\\Users\\rudu0916\\Desktop\\SDN_NFV\\du\\DU vCPE BidSheet v19_1.xlsm'
CM91_WBSpath = 'C:\\Users\\rudu0916\\Desktop\\SDN_NFV\\HOM\\CM\\CM R9.1 Proposal\\Worksheet in CM91_proposal_v1_RD.xlsx'

names = ['Story',
         'FR',
         'Feature',
         'Service',
        'Analysis LOE wbs',
        'Design LOE wbs', 
        'Build LOE wbs',
        'Bug Fixing LOE wbs',
        'QA LOE wbs',
        'Tech LOE wbs']
subset = ['FR']

CM91_WBS = wbs.loadWBS(CM91_WBSpath,names,subset,index_col=None,sheetname='Product LOE',parse_cols='A:D,F:K', how='any')

CM91_WBS[['Work Category', 'Module']] = du_vCPE_WBS[['Work Category', 'Module']].fillna('General')

CM91_WBS[['FR','Feature','Tech LOE wbs']].groupby(['FR','Feature']).sum().to_excel('C:\\Users\\rudu0916\\Desktop\\SDN_NFV\\HOM\\CM\\CM R9.1 Proposal\\Fwatures.xlsx', merge_cells=False)

CM91_WBS_summary = du_vCPE_WBS[['Phase', 'Work Category', 'Module', 'Specialization', 'Analysis LOE wbs', 'Dev LOE wbs', 'QA LOE wbs']].groupby(by=['Phase', 'Work Category', 'Module', 'Specialization']).sum()

CM91_WBS_summary.to_excel('C:\\Users\\rudu0916\\Desktop\\SDN_NFV\\du\\DU vCPE BidSheet v19_1 summary.xlsx', merge_cells=False)

CM91_WBS.columns


# 'HL Estimate, md'     customfield_23121 - calculated, cannot be updated
# 'Analysis LOE, md'    customfield_23422
# 'Design LOE, md'      customfield_23423
# 'Build LOE, md'       customfield_23424
# 'Bug Fixing LOE, md'  customfield_23425
# 'QA LOE, md'          customfield_23426

wbs.allWBStoHLEstimate(jira, UUI_WBS)


names = [#'Scope of Work',
         'Customer',
         'Bucket',
         'Team',
        'Analysis LOE wbs',
        'Design LOE wbs', 
        'Build LOE wbs',
        'Bug Fixing LOE wbs',
        'QA LOE wbs',
        'Tech LOE wbs']
subset = [#'Scope of Work',
         'Bucket',
         'Tech LOE wbs']

TFA9_WBS = wbs.loadWBS(TFA9_WBSpath,names,subset,index_col=[4,0], sheetname='Project and Product LOE',parse_cols='A,C:E,G,I:N',how='any')
TFA9_RP = wbs.loadRP(TFA9_WBSpath,
            skiprows=1,
            index_col=3,
            sheetname='Resource Plan',
            parse_cols='A:I')


TFA9_WBS = TFA9_WBS[['Analysis LOE wbs',
        'Design LOE wbs', 
        'Build LOE wbs',
        'Bug Fixing LOE wbs',
        'QA LOE wbs',
        'Tech LOE wbs']].groupby(level=0).sum()
wbs.allWBStoHLEstimate(jira, TFA9_WBS, force = True)


TFA9_WBS = wbs.getGross(TFA9_WBS, TFA9_RP)


TFA9_RP_Teams_summary = TFA9_RP[['Total','FTE','Phase','Team']].groupby(by=['Phase','Team']).sum()
TFA9_RP_Teams_summary['Total'] = TFA9_RP_Teams_summary['Total']/8
TFA9_RP_summary = TFA9_RP[['Total','FTE','Phase']].groupby(by='Phase').sum()
TFA9_RP_summary['Total'] = TFA9_RP_summary['Total']/8

writer = pd.ExcelWriter('C:\\Users\\rudu0916\\Desktop\\SDN_NFV\\HOM\\TFA&OTM\\TFA Proposals 9.0\\HOMTFAR90_RP_summary.xlsx',
                        engine='xlsxwriter')

TFA9_RP_Teams_summary.to_excel(writer,
                        merge_cells=False,
                        sheet_name = 'Resource Plan')
    
TFA9_RP_summary.to_excel(writer,
                    merge_cells=False,
                    sheet_name = 'Resource Plan',
                    startrow = 11)

writer.save()

#TFA9_RP_summary.loc['MGMT']['Total']

#TFA9_WBS[['Impl LOE wbs','Team','Bucket']].groupby(by=['Team','Bucket']).sum()
TFA9_WBS_summary = TFA9_WBS.sum()
Contingency = (TFA9_WBS_summary['Tech LOE wbs'] + TFA9_RP_summary.loc['MGMT']['Total']) * 0.2
Gross_LOE =  TFA9_WBS_summary['Tech LOE wbs'] + TFA9_RP_summary.loc['MGMT']['Total'] + Contingency

#TFA9_WBS['Gross_LOE'].sum()

TFA9_WBS[['Tech LOE wbs', 'Management LOE', 'Contingency', 'Gross LOE','Bucket']].groupby(by='Bucket').sum()
TFA9_WBS[['Tech LOE wbs', 'Dev LOE wbs','Team']].groupby(by='Team').sum()


TFA9_RP_summary['WBS total'] =  TFA9_RP_summary.apply(lambda x: wbs.getWBStotal(x.name, TFA9_WBS_summary), axis=1)


wbs.writeSummary(TFA9_WBS, TFA9_RP, 'C:\\Users\\rudu0916\\Desktop\\SDN_NFV\\HOM\\TFA&OTM\\TFA Proposals 9.0\\Product_TFA-R9.0v22_rd_summary.xlsx')



VDOC_names = [#'Scope of Work',
             #'FR',
             'Team',
             'Customer',
             'Bucket',
             'Analysis LOE wbs',
             'Design LOE wbs', 
             'Build LOE wbs',
             'Bug Fixing LOE wbs',
             'QA LOE wbs',
             'Tech LOE wbs']
VDOC_subset = [#'Scope of Work',
            'Bucket',
            'Tech LOE wbs']

VDOCR5_WBS = wbs.loadWBS(VDOCR4_WBSpath,VDOC_names,VDOC_subset,skiprows=2,index_col=[1,2,0],sheetname = 'Product LOE',parse_cols='A:F,L:Q',how='any')
VDOCR5_RP = wbs.loadRP(VDOCR5_WBSpath,
            skiprows=2,
            index_col=3,
            sheetname='Resource Plan',
            parse_cols='A:M')

VDOCR4_WBS = VDOCR4_WBS[['Analysis LOE wbs',
        'Design LOE wbs', 
        'Build LOE wbs',
        'Bug Fixing LOE wbs',
        'QA LOE wbs',
        'Tech LOE wbs']].groupby(level=0).sum()

wbs.allWBStoHLEstimate(jira, VDOCR4_WBS, force = True)

VDOCR4_WBS = wbs.getGross(VDOCR4_WBS, VDOCR4_RP)


VDOCR5_RP_Teams_summary = VDOCR5_RP[['Total','FTE','Phase','Team']].groupby(by=['Phase','Team']).sum()
VDOCR5_RP_Teams_summary['Total'] = VDOCR5_RP_Teams_summary['Total']/8
VDOCR5_RP_summary = VDOCR5_RP[['Total','FTE','Phase']].groupby(by='Phase').sum()
VDOCR5_RP_summary['Total'] = VDOCR5_RP_summary['Total']/8

VDOCR5_RP_summary.sum()

#UI9_RP_summary.loc['MGMT']['Total']

#UI9_WBS[['Dev LOE wbs','Team']].groupby(by='Team').sum()
VDOCR4_WBS_summary = VDOCR4_WBS.sum()
#Contingency = (VDOCR4_WBS_summary['Tech LOE wbs'] + VDOCR4_RP_summary.loc['MGMT']['Total']) * 0.2
#Gross_LOE =  VDOCR4_WBS_summary['Tech LOE wbs'] + VDOCR4_RP_summary.loc['MGMT']['Total'] + Contingency

#UI9_WBS['Gross_LOE'].sum()

#VDOCR4_WBS[['Tech LOE wbs', 'Management LOE', 'Contingency', 'Gross LOE','Customer type']].groupby(by='Customer type').sum()


VDOCR4_RP_summary['WBS total'] =  VDOCR4_RP_summary.apply(lambda x: wbs.getWBStotal(x.name, VDOCR4_WBS_summary), axis=1)


wbs.writeSummary(VDOCR4_WBS, VDOCR4_RP, 'C:\\Users\\rudu0916\\Desktop\\SDN_NFV\\AVP\\R4 Proposal\\VDOC_R4_WBS_v2_summary.xlsx')




for x in VDOCR4_WBS.iterrows():
    if x[1]['Tech LOE wbs'] == 0:
        print(x[0][0] + ' skipped')
    else:
        print(x[0][0])
        if wbs.WBStoDecription(jira, x, tag = 'VDOCR4_WBS'):
            print('Updated')




        fields = jira.fields()
        [x['id'] for x in fields if x['name'] == 'Description'][0]







#Need to load all columns to have Position with teams division
#Delete unnecessary columns by title. Create column names from titles row
#Identify sprints start and end dates, indices of sprint starts and ends
#Group sum by team

path = 'C:\\Users\\rudu0916\\Desktop\\SDN_NFV\\AVP\\SM\\R4\\'
file = 'NC.SDN.PRD.VDOCSMR4_RP_V3_07.11.2017.xlsm'


sprints = wbs.getSprints(path+file, parse_cols='A,D,E:Q', sprintRow=13)

VDOCR4_RP = wbs.loadRP(path+file,
            skiprows=2,
            index_col=3,
            sheetname='Resource Plan',
            parse_cols='A:Q')




for (sprint, params) in sprints.items():
    VDOCR4_RP[sprint] = VDOCR4_RP[list(range(params['start_id'],params['end_id']+1))].apply(sum,axis=1)

phases = ['DEV']

VDOCR4_RP.query('Phase in ' + str(phases))[['Team']+list(sprints.keys())].groupby(by='Team').sum()


sprint_capacity = 0
sprint_start = 0
last_sprint_start = 0

for week in WBS.columns:
    sprint = WBS[week][sprintRow]
    if str(sprint) != 'nan':
        sprint_name = sprint
        sptint_start_id = WBS[week][0]
        sprint_capacity

        if sprint_start != 0:
            sprint_start = last_sprint_start
        else:
            sprint_start = week
            last_sprint_start = week
            sptint_start_id = WBS[week][0]


UI9_RP = wbs.loadRP(UI9_WBSpath,
            skiprows=1,
            index_col=3,
            sheetname='Resource Plan',
            parse_cols='A:N')
