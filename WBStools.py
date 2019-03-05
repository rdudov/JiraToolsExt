import pandas as pd
import re
import pandas.io.formats.excel
import common as cmn
import datetime
from xlsxwriter.utility import xl_rowcol_to_cell

def loadWBS(WBSpath,
            names,
            subset,
            skiprows=2,
            index_col=1,
            sheetname='Product LOE',
            parse_cols='A:E,G,I:O',
            how='all'):

    WBS = pd.read_excel(WBSpath,
                        sheetname,
                        header=0,
                        skiprows=skiprows,
                        index_col=index_col,
                        names=names,
                        parse_cols=parse_cols)

    WBS.dropna(subset=subset,inplace=True,how=how)
    
    WBS['Analysis LOE wbs'].fillna(0, inplace = True)
    WBS['Design LOE wbs'].fillna(0, inplace = True)
    WBS['Build LOE wbs'].fillna(0, inplace = True)
    WBS['Bug Fixing LOE wbs'].fillna(0, inplace = True)
    WBS['QA LOE wbs'].fillna(0, inplace = True)

    WBS['Impl LOE wbs'] = WBS.apply(lambda row: row['Design LOE wbs'] + row['Build LOE wbs'],axis=1)
    WBS['Dev LOE wbs'] = WBS.apply(lambda x: x['Impl LOE wbs']+x['Bug Fixing LOE wbs'], axis=1)

    return WBS

def getPhase(position):
    m = re.findall('\[(\w+)\]',position)
    if len(m) == 0:
        return ''
    else:
        return m[-1]

def getTeam(position):
    m = re.findall('\[(.+?)\]',position)
    if len(m) == 0:
        return ''
    else:
        return m[0]

def getWBStotal(RPphase, WBS_summary):
    if RPphase == 'BA':
        return WBS_summary['Analysis LOE wbs']
    elif RPphase == 'DEV':
        return WBS_summary['Dev LOE wbs']
    elif RPphase == 'QA':
        return WBS_summary['QA LOE wbs']
    else:
        return 0

def getGrossParts(WBS_item, Total_Tech_LOE, Total_Management_LOE, Total_Contingency, Total_Gross_LOE):
    rate = WBS_item['Tech LOE wbs'] / Total_Tech_LOE
    Gross_LOE = round(rate * Total_Gross_LOE,1)
    Management_LOE = round(rate * Total_Management_LOE,1)
    Contingency = round(rate * Total_Contingency,1)
    return pd.Series({'Management LOE' : Management_LOE,
                      'Contingency' : Contingency,
                      'Gross LOE' : Gross_LOE})

def getGross(WBS, RP):
    
    RP_summary = RP[['Total','Phase']].groupby(by='Phase').sum()/8

    Total_Tech_LOE = WBS['Tech LOE wbs'].sum()
    Total_Management_LOE = RP_summary.loc['MGMT']['Total']
    Total_Contingency = (Total_Tech_LOE + Total_Management_LOE) * 0.2
    Total_Gross_LOE =  Total_Tech_LOE + Total_Management_LOE + Total_Contingency

    return WBS.join(WBS.apply(lambda x: getGrossParts(x,
                                                     Total_Tech_LOE,
                                                     Total_Management_LOE,
                                                     Total_Contingency,
                                                     Total_Gross_LOE), axis=1))

def loadRP(RPpath,
            skiprows=2,
            index_col=2,
            sheetname='Resource Plan',
            parse_cols='A:P',
            weeks=0):

    RP = pd.read_excel(RPpath,
                        sheetname,
                        header=0,
                        skiprows=skiprows,
                        index_col=index_col,
                        parse_cols=parse_cols,
                        weeks=weeks)

    RP.dropna(subset=['Position','Role'], inplace=True)
    RP['Phase'] = RP.apply(lambda x: getPhase(x['Position']), axis=1)
    RP['Team'] = RP.apply(lambda x: getTeam(x['Position']), axis=1)

    if weeks > 0:
        weeks_index = list(range(1,weeks+1))
    else:
        weeks_index = list(filter(lambda x: isinstance(x, int),RP.columns.values))

    RP['Total'] = RP[weeks_index].sum(axis=1)
    RP['FTE'] = RP['Total'] / (len(weeks_index) * 37)
    return RP


def writeSummary(WBS, RP, path):
    format_all = {'font_size': 9,
                'bottom': 1,
                'right': 1,
                'top': 1,
                'left': 1}
    format_long = {**format_all, **{'text_wrap': True}}
    header_fmt = {**format_all, **{'bold': True,
                                    'text_wrap': True,
                                    'align': 'center',
                                    'valign': 'vcenter'}}
    effort_fmt = {**format_all, **{'num_format': '0.0'}}

    WBS_columns = ['Customer',
                 'Bucket',
                 'Team',
                'Analysis LOE wbs',
                'Design LOE wbs', 
                'Build LOE wbs',
                'Bug Fixing LOE wbs',
                'QA LOE wbs',
                'Tech LOE wbs', 
                        'Gross LOE']
    WBS_labels = ['Customer',
                 'Bucket',
                 'Team',
                   'Analysis LOE wbs',
                    'Design LOE wbs', 
                    'Build LOE wbs',
                    'Bug Fixing LOE wbs',
                    'QA LOE wbs',
                    'Tech LOE wbs', 
                        'Gross LOE']
    Summary_columns = ['Tech LOE wbs',
                        'Management LOE',
                        'Contingency', 
                        'Gross LOE']
    Summary_labels = ['Tech LOE, md\nA+D+B+BF+QA',
                       'Management + Overheads, md',
                        'Contingency 20%, md', 
                        'Total LOE, md']

    WBS_Summary = WBS[['Bucket'] + Summary_columns].groupby(by='Bucket').sum()

    writer = pd.ExcelWriter(path, engine='xlsxwriter')

    pd.io.formats.excel.header_style = None

    WBS.to_excel(writer,
                index_label = ['MANOFR', 'Scope of Work'],
                merge_cells=False,
                sheet_name = 'Scope',
                columns = WBS_columns,
                header = WBS_labels)

    workbook = writer.book
    
    wb_format_all = workbook.add_format(format_all)
    wb_format_long = workbook.add_format(format_long)
    wb_header_fmt = workbook.add_format(header_fmt)
    wb_effort_fmt = workbook.add_format(effort_fmt)

    worksheet = writer.sheets['Scope']

    worksheet.set_column('A:A', 11, wb_format_all)
    worksheet.set_column('B:B', 40, wb_format_long)
    worksheet.set_column('C:D', 15, wb_format_all)
    worksheet.set_column('E:E', 7, wb_format_all)
    worksheet.set_column('F:P', 8, wb_effort_fmt)

    worksheet.set_row(0, None, wb_header_fmt)

    WBS_Summary.to_excel(writer,
                    index_label = ['Bucket'],
                    merge_cells=False,
                    sheet_name = 'Summary',
                    columns = Summary_columns,
                    header = Summary_labels)
    
    worksheet = writer.sheets['Summary']

    total_effort_fmt = workbook.add_format({**effort_fmt, **{'bold': True, 'num_format': '0'}})
    effort_sum_fmt = workbook.add_format({**effort_fmt, **{'num_format': '0'}})

    worksheet.set_column('A:A', 25, wb_format_all)
    worksheet.set_column('B:D', 12, effort_sum_fmt)
    worksheet.set_column('E:E', 12, total_effort_fmt)

    worksheet.set_row(0, None, wb_header_fmt)

    number_rows = len(WBS_Summary.index)
    number_columns = len(WBS_Summary.columns)

    worksheet.write_string(number_rows+1, 0, 'Total for All activities', total_effort_fmt)

    for column in range(1, number_columns + 1):
        worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, column),
                                "=SUM({:s}:{:s})".format(xl_rowcol_to_cell(1, column), xl_rowcol_to_cell(number_rows, column)),
                                total_effort_fmt)

    RP_Teams_summary = RP[['Total','FTE', 'Phase','Team']].groupby(by=['Phase','Team']).sum()
    RP_Teams_summary['Total'] = RP_Teams_summary['Total']/8
    RP_summary = RP[['Total','FTE','Phase']].groupby(by='Phase').sum()
    RP_summary['Total'] = RP_summary['Total']/8
    WBS_totals = WBS.sum()
    RP_summary['WBS total'] =  RP_summary.apply(lambda x: getWBStotal(x.name, WBS_totals), axis=1)
    
    RP_Teams_summary.to_excel(writer,
                            merge_cells=False,
                            sheet_name = 'Resources')
    
    RP_summary.to_excel(writer,
                        merge_cells=False,
                        sheet_name = 'Resources',
                        startcol = 6)

    writer.save()

    return

def WBStoDecription(jira, WBSitem, tag = 'WBS'):
        issue = jira.issue(WBSitem[0])
        description = issue.fields.description

        WBSdata = '{"HL Estimate, md" : ' + str(WBSitem[1]['Tech LOE wbs']) + \
                    ',"Analysis LOE, md" : ' + str(WBSitem[1]['Analysis LOE wbs']) +\
                    ',"Design LOE, md" : ' + str(WBSitem[1]['Design LOE wbs']) +\
                    ',"Build LOE, md" : ' + str(WBSitem[1]['Build LOE wbs']) +\
                    ',"Bug Fixing LOE, md" : ' + str(WBSitem[1]['Bug Fixing LOE wbs']) +\
                    ',"QA LOE, md" : ' + str(WBSitem[1]['QA LOE wbs']) + '}'

        if not description is None:
            m = re.search('#' + tag + '(\{.*\})',description)
            if not m is None:
                print('Already filled: ' + m.group(0))
                return False
            description = description + '\n\n#' + tag + WBSdata
        else:
            description = '#' + tag + WBSdata

        issue.update(description = description,notify=False)

        return True

# 'HL Estimate, md'     customfield_23121 - calculated, cannot be updated
# 'Analysis LOE, md'    customfield_23422
# 'Design LOE, md'      customfield_23423
# 'Build LOE, md'       customfield_23424
# 'Bug Fixing LOE, md'  customfield_23425
# 'QA LOE, md'          customfield_23426
def WBStoHLEstimate(issue, WBSitem):
    issue.update(customfield_23422=round(WBSitem[1]['Analysis LOE wbs'],1),
                 customfield_23423=round(WBSitem[1]['Design LOE wbs'],1),
                 customfield_23424=round(WBSitem[1]['Build LOE wbs'],1),
                 customfield_23425=round(WBSitem[1]['Bug Fixing LOE wbs'],1),
                 customfield_23426=round(WBSitem[1]['QA LOE wbs'],1),notify=False)
    return

def allWBStoHLEstimate(jira, WBS, force=False):
    for x in WBS.iterrows():
        if str(x[0]) == 'nan':
            print('skipped - no FR')
        else:
            print(x[0])
            try:
                issue = jira.issue(x[0])
                total_jira = round(cmn.nvl(issue.fields.customfield_23121,0),1)
                analysis_jira = round(cmn.nvl(issue.fields.customfield_23422,0),1)
                design_jira = round(cmn.nvl(issue.fields.customfield_23423,0),1)
                build_jira = round(cmn.nvl(issue.fields.customfield_23424,0),1)
                bugfix_jira = round(cmn.nvl(issue.fields.customfield_23425,0),1)
                qa_jira = round(cmn.nvl(issue.fields.customfield_23426,0),1)

                if abs(total_jira - round(x[1]['Tech LOE wbs'],1)) > 0.2:
                    print('Different estimations exist:')
                    print('HL Estimate, md      ' + str(total_jira) +'    ' + str(round(x[1]['Tech LOE wbs'],1)))
                    print('Analysis LOE, md     ' + str(analysis_jira) + '    ' + str(round(x[1]['Analysis LOE wbs'],1)))
                    print('Design LOE, md       ' + str(design_jira) + '    ' + str(round(x[1]['Design LOE wbs'],1)))
                    print('Build LOE, md        ' + str(build_jira) + '    ' + str(round(x[1]['Build LOE wbs'],1)))
                    print('Bug Fixing LOE, md   ' + str(bugfix_jira) + '    ' + str(round(x[1]['Bug Fixing LOE wbs'],1)))
                    print('QA LOE, md           ' + str(qa_jira) + '    ' + str(round(x[1]['QA LOE wbs'],1)))
                    
                    if force:
                        WBStoHLEstimate(issue, x)
                        print('Updated with force')

                elif (total_jira == 0) and (x[1]['Tech LOE wbs'] != 0):
                    WBStoHLEstimate(issue, x)
                    print('Updated')

                else:
                    print('Already filled')

            except Exception as e:
                print(e)
                continue
    return

def getSprints(RPfile,
               parse_cols,
               sprintRow,
               skiprows=1,
               sheetname='Resource Plan'):

    WBS = pd.read_excel(RPfile,
                        skiprows=skiprows,
                        sheetname=sheetname,
                        usecols=parse_cols) 

    weeksRow = 0

    weeksStartCol = None

    for sprint in WBS.iloc[sprintRow].tolist():
        if weeksStartCol is None:
            if str(sprint) != 'nan':
                currCol = WBS.iloc[sprintRow].tolist().index(sprint)
                weeksStartCol = currCol
                sprints = {sprint : {'start_id' : int(WBS.iloc[weeksRow].tolist()[currCol]),
                                     'start_date' : WBS.columns[currCol]}}
                lastSprint = sprint
        elif (str(sprint) != 'nan') and (lastSprint != sprint):
            currCol = WBS.iloc[sprintRow].tolist().index(sprint)
            sprints[lastSprint] = {**sprints[lastSprint], **{'end_id' : int(WBS.iloc[weeksRow].tolist()[currCol-1]),
                                                             'end_date' : WBS.columns[currCol-1] + datetime.timedelta(days=6)}}
            sprints = {**sprints, **{sprint : {'start_id' : int(WBS.iloc[weeksRow].tolist()[currCol]),
                                               'start_date' : WBS.columns[currCol]}}}
            lastSprint = sprint

    sprints[lastSprint] = {**sprints[lastSprint], **{'end_id' : int(WBS.iloc[weeksRow].tolist()[-1]),
                                                     'end_date' : WBS.columns[-1] + datetime.timedelta(days=6)}}


    return sprints