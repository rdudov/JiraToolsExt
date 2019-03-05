import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell
import pandas.io.formats.excel
import jiraConnector as jc
import WBStools as wbs
import numpy as np
import common as cmn

def original_estimate(row):
    return max(row['Original Estimate'], row['Time Spent'] + row['Estimate'])

def remaining_estimate(row):
    if (row['Status Mapped']=='Done') | (row['Status Mapped']=='In QA'):
        return 0
    else:
       return row['Remaining Estimate']

def calc_subtasks(WorkItem, WorkItems, subtasks):

    WorkItem_row = WorkItems[WorkItems.index.get_level_values(1) == WorkItem].iloc[0]
    
    workitem_estimate = original_estimate(WorkItem_row)
    workitem_fact = WorkItem_row['Time Spent']
    total_fact = WorkItem_row['Σ Time Spent']
    
    workitem_done_cnt = 0
    workitem_done_estimate = 0
    workitem_implemented_cnt = 0
    workitem_implemented_estimate = 0
    workitem_remaining_estimate = 0

    if workitem_estimate > 0:
        workitem_cnt = 1
        if WorkItem_row['Status Mapped'] == 'Done':
            workitem_done_cnt = 1
            workitem_implemented_cnt = 1
            workitem_implemented_estimate = workitem_estimate
            workitem_done_estimate = workitem_estimate
        elif WorkItem_row['Status Mapped'] == 'In QA':
            workitem_implemented_cnt = 1
            workitem_implemented_estimate = workitem_estimate
        else:
            workitem_remaining_estimate = WorkItem_row['Remaining Estimate']
    else:
        workitem_cnt = 0
        
    subtasks_c = subtasks[subtasks.index.get_level_values(1).isin([WorkItem])]
    cnt_subtasks = len(subtasks_c)

    cnt_total = cnt_subtasks + workitem_cnt

    cnt_nonqa_subtasks = len(subtasks_c[subtasks_c['Step']!='QA'])
    cnt_nonqa = cnt_nonqa_subtasks + workitem_cnt

    cnt_qa = cnt_total-cnt_nonqa

    cnt_done_subtasks = len(subtasks_c[subtasks_c['Status Mapped']=='Done'])
    cnt_done = cnt_done_subtasks + workitem_done_cnt

    cnt_done_qa = len(subtasks_c[(subtasks_c['Step']=='QA') & (subtasks_c['Status Mapped']=='Done')])

    cnt_dev_implemented_subtasks = len(subtasks_c[(subtasks_c['Step']!='QA') & subtasks_c['Status Mapped'].isin(['In QA', 'Done'])])
    cnt_dev_implemented = cnt_dev_implemented_subtasks + workitem_implemented_cnt

    if cnt_subtasks > 0:
        total_estimate = sum(subtasks_c.apply(original_estimate,axis=1)) + workitem_estimate
        total_remaining = sum(subtasks_c.apply(remaining_estimate,axis=1)) + workitem_remaining_estimate
    else:
        total_estimate = workitem_estimate
        total_remaining = workitem_remaining_estimate

    if cnt_nonqa_subtasks > 0:
        nonqa_estimate = sum(subtasks_c[subtasks_c['Step']!='QA'].apply(original_estimate,axis=1)) + workitem_estimate
        nonqa_fact = sum(subtasks_c[subtasks_c['Step']!='QA'].apply(lambda row: row['Time Spent'],axis=1)) + workitem_fact
        nonqa_remaining = sum(subtasks_c[subtasks_c['Step']!='QA'].apply(remaining_estimate,axis=1)) + workitem_remaining_estimate
    else:
        nonqa_estimate = workitem_estimate
        nonqa_fact = workitem_fact
        nonqa_remaining = workitem_remaining_estimate

    qa_estimate = total_estimate - nonqa_estimate
    qa_fact = total_fact - nonqa_fact
    qa_remaining = total_remaining - nonqa_remaining

    if cnt_done_subtasks > 0:
        done_estimate = sum(subtasks_c[subtasks_c['Status Mapped']=='Done'].apply(original_estimate,axis=1)) + workitem_done_estimate
    else:
        done_estimate = workitem_done_estimate

    if cnt_done_qa > 0:
        done_qa_estimate = sum(subtasks_c[(subtasks_c['Step']=='QA') & (subtasks_c['Status Mapped']=='Done')].apply(original_estimate,axis=1))
    else:
        done_qa_estimate = 0

    if cnt_dev_implemented_subtasks > 0:
        implemented_dev_estimate = sum(subtasks_c[(subtasks_c['Step']!='QA') & subtasks_c['Status Mapped'].isin(['In QA', 'Done'])].apply(original_estimate,axis=1)) + workitem_implemented_estimate
    else:
        implemented_dev_estimate = workitem_implemented_estimate

    if total_estimate > 0:
        total_rate = done_estimate / total_estimate
        total_spent_rate = total_fact / total_estimate
    else:
        if cnt_done > 0:
            total_rate = cnt_done / cnt_total
        else:
            total_rate = 0
        total_spent_rate = 0

    if nonqa_estimate > 0:
        nonqa_rate = implemented_dev_estimate / nonqa_estimate
        nonqa_spent_rate = nonqa_fact / nonqa_estimate
    else:
        if cnt_dev_implemented > 0:
            nonqa_rate = cnt_dev_implemented / cnt_nonqa
        else:
            nonqa_rate = 0
        nonqa_spent_rate = 0
        
    if qa_estimate > 0:
        qa_rate = done_qa_estimate / qa_estimate
        qa_spent_rate = qa_fact / qa_estimate
    else:
        if cnt_done_qa > 0:
            qa_rate = cnt_done_qa / cnt_qa
        else:
            qa_rate = 0
        qa_spent_rate = 0
    
    if WorkItem_row['Σ Original Estimate'] > 0:
        total_fcast_plan_rate = total_estimate / WorkItem_row['Σ Original Estimate']
    else:
        total_fcast_plan_rate = 0
    
    if (cnt_total == 0) | (total_estimate == 0):
        if WorkItem_row['Status Mapped'] == 'Done':
            total_rate = 1
            nonqa_rate = 1
            qa_rate = 1
        elif WorkItem_row['Status Mapped'] == 'In QA':
            total_rate = 0
            nonqa_rate = 1
            qa_rate = 0

    return pd.Series({'cnt_total': cnt_total,
                     'cnt_nonqa': cnt_nonqa,
                     'cnt_qa': cnt_qa,
                     'cnt_done': cnt_done,
                     'cnt_done_qa': cnt_done_qa,
                     'cnt_dev_implemented': cnt_dev_implemented,
                     'total_estimate': total_estimate,
                     'nonqa_estimate': nonqa_estimate,
                     'qa_estimate': qa_estimate,
                     'total_fcast_plan_rate': total_fcast_plan_rate,
                     'workitem_fact': workitem_fact,
                     'nonqa_fact': nonqa_fact,
                     'qa_fact': qa_fact,
                     'total_remaining': total_remaining,
                     'nonqa_remaining': nonqa_remaining,
                     'qa_remaining': qa_remaining,
                     'total_spent_rate': total_spent_rate,
                     'nonqa_spent_rate': nonqa_spent_rate,
                     'qa_spent_rate': qa_spent_rate,
                     'done_estimate': done_estimate,
                     'done_qa_estimate': done_qa_estimate,
                     'implemented_dev_estimate': implemented_dev_estimate,
                     'total_rate': total_rate,
                     'nonqa_rate': nonqa_rate,
                     'qa_rate': qa_rate})

def format_WorkItems(workbook,
                    worksheet,
                    number_rows,
                    p_format_all,
                    p_header_fmt,
                    p_effort_fmt,
                    p_percent_fmt,
                    p_format_alert,
                    p_format_warn,
                    p_format_impl,
                    p_format_done,
                    format_bar_impl,
                    format_bar_done):
    
    format_all = workbook.add_format(p_format_all)
    header_fmt = workbook.add_format(p_header_fmt)
    effort_fmt = workbook.add_format(p_effort_fmt)
    percent_fmt = workbook.add_format(p_percent_fmt)

    worksheet.set_column('A:B', 11, format_all)
    worksheet.set_column('C:C', 40, format_all)
    worksheet.set_column('D:D', 11, format_all)
    worksheet.set_column('E:E', 7, format_all)
    worksheet.set_column('F:F', 15, format_all)
    worksheet.set_column('G:G', 10, format_all)
    worksheet.set_column('H:K', 8, effort_fmt)
    worksheet.set_column('L:N', 8, percent_fmt)
    worksheet.set_column('O:Q', 8, effort_fmt)
    worksheet.set_column('R:S', 8, percent_fmt)
    worksheet.set_column('T:V', 8, effort_fmt)
    worksheet.set_column('W:X', 8, percent_fmt)
    worksheet.set_column('Y:AN',8, format_all)

    worksheet.set_row(0, None, header_fmt)
    
    format_alert = workbook.add_format(p_format_alert)
    format_warn = workbook.add_format(p_format_warn)
    
    format_impl = workbook.add_format(p_format_impl)
    format_done = workbook.add_format(p_format_done)
    
    #Fact exceeds plan (forecast rate * spent rate) for over 20%
    worksheet.conditional_format('J2:J{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=($L2*$M2)>=1.2', 'format': format_alert})
    
    #WorkItem is completed
    worksheet.conditional_format('A2:AN{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=$N2=1', 'format': format_done})
    
    #Development is completed
    worksheet.conditional_format('A2:AV{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=$S2=1', 'format': format_impl})
    
    #Progress pars
    worksheet.conditional_format('N2:N{}'.format(number_rows+2), format_bar_done)
    worksheet.conditional_format('S2:S{}'.format(number_rows+2), format_bar_impl)
    worksheet.conditional_format('X2:X{}'.format(number_rows+2), format_bar_impl)
    
    #Gained effort exceeds gained progress for over then 20%
    worksheet.conditional_format('R2:R{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=($R2-$S2)>=0.2', 'format': format_warn})

    #Forecast exceeds plan for over than 20%
    worksheet.conditional_format('L2:L{}'.format(number_rows+2),
                                    {'type': 'cell',
                                    'criteria': '>=',
                                    'value': 1.2,
                                    'format': format_alert})

    total_effort_fmt = workbook.add_format({**p_effort_fmt, **{'bold': True, 'top': 6}}) 
    total_percent_fmt = workbook.add_format({**p_percent_fmt, **{'bold': True, 'top': 6}})

    worksheet.write_string(number_rows+1, 5, "Total",total_effort_fmt)

    for column in [7, 8, 9, 10, 14, 15, 16, 19, 20, 21]:
        worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, column),
                                "=SUM({:s}:{:s})".format(xl_rowcol_to_cell(1, column), xl_rowcol_to_cell(number_rows, column)),
                                total_effort_fmt)
        
    #Total fcast/plan
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 11),
                            "={:s}/{:s}".format(xl_rowcol_to_cell(number_rows+1, 8),
                                                xl_rowcol_to_cell(number_rows+1, 7)),
                            total_percent_fmt)
    #Total fact/fcast
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 12),
                            "={:s}/{:s}".format(xl_rowcol_to_cell(number_rows+1, 9),
                                                xl_rowcol_to_cell(number_rows+1, 8)),
                            total_percent_fmt)
    #Done, %
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 13),
                            "=SUM({:s}:{:s})/SUM({:s}:{:s})".format(xl_rowcol_to_cell(1, 35),
                                                                    xl_rowcol_to_cell(number_rows, 35),
                                                                    xl_rowcol_to_cell(1, 8),
                                                                    xl_rowcol_to_cell(number_rows, 8)),
                            total_percent_fmt)

    #Impl fact/fcast
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 17),
                            "={:s}/{:s}".format(xl_rowcol_to_cell(number_rows+1, 15),
                                                xl_rowcol_to_cell(number_rows+1, 14)),
                            total_percent_fmt)
    #Impl, %
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 18),
                            "=SUM({:s}:{:s})/SUM({:s}:{:s})".format(xl_rowcol_to_cell(1, 37),
                                                                    xl_rowcol_to_cell(number_rows, 37),
                                                                    xl_rowcol_to_cell(1, 14),
                                                                    xl_rowcol_to_cell(number_rows, 14)),
                            total_percent_fmt)

    #QA fact/fcast
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 22),
                            "={:s}/{:s}".format(xl_rowcol_to_cell(number_rows+1, 20),
                                                xl_rowcol_to_cell(number_rows+1, 19)),
                            total_percent_fmt)
    #QA, %
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 23),
                            "=SUM({:s}:{:s})/SUM({:s}:{:s})".format(xl_rowcol_to_cell(1, 36),
                                                                    xl_rowcol_to_cell(number_rows, 36),
                                                                    xl_rowcol_to_cell(1, 19),
                                                                    xl_rowcol_to_cell(number_rows, 19)),
                            total_percent_fmt)
        
    worksheet.autofilter(0,0,number_rows+1,39)

    

def format_subtasks(workbook,
                   worksheet,
                   subtasks,
                    p_format_all,
                    p_header_fmt,
                    p_effort_fmt,
                    p_format_impl,
                    p_format_done):
    
    number_rows = len(subtasks.index)
    format_all = workbook.add_format(p_format_all)
    header_fmt = workbook.add_format(p_header_fmt)
    effort_fmt = workbook.add_format(p_effort_fmt)

    worksheet.set_column('A:C', 11, format_all)
    worksheet.set_column('D:D', 40, format_all)
    worksheet.set_column('E:E', 11, format_all)
    worksheet.set_column('F:F', 7, format_all)
    worksheet.set_column('G:G', 15, format_all)
    worksheet.set_column('H:H', 10, format_all)
    worksheet.set_column('I:K', 8, effort_fmt)
    worksheet.set_column('L:P',8, format_all)

    worksheet.set_row(0, None, header_fmt)
    
    format_impl = workbook.add_format(p_format_impl)
    format_done = workbook.add_format(p_format_done)
    
    #Resolved
    worksheet.conditional_format('A2:P{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=$P2="In QA"', 'format': format_impl})

    #Closed
    worksheet.conditional_format('A2:P{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=$P2="Done"', 'format': format_done})

    worksheet.autofilter(0,0,number_rows+1,len(subtasks.columns)+1)

    return


def write_report(path,
                WorkItems,
                subtasks,
                WorkItems_columns,
                WorkItems_labels,
                subtasks_columns,
                format_all,
                header_fmt,
                effort_fmt,
                percent_fmt,
                format_alert,
                format_warn,
                format_impl,
                format_done,
                format_bar_impl,
                format_bar_done):
    
    writer = pd.ExcelWriter(path, engine='xlsxwriter')

    pd.io.formats.excel.header_style = None

    workbook = writer.book

    WorkItems.to_excel(writer,
                    index_label = ['Epic', 'Key'],
                    merge_cells=False,
                    freeze_panes=(1,2),
                    sheet_name = 'WorkItems',
                    columns = WorkItems_columns,
                    header = WorkItems_labels)
    
    format_WorkItems(workbook, writer.sheets['WorkItems'],
                    len(WorkItems.index),
                    format_all,
                    header_fmt,
                    effort_fmt,
                    percent_fmt,
                    format_alert,
                    format_warn,
                    format_impl,
                    format_done,
                     format_bar_impl,
                     format_bar_done)

    subtasks.to_excel(writer,
                      index_label = ['Epic', 'WorkItem','Key'],
                      merge_cells=False,
                      freeze_panes=(1,3),
                      sheet_name = 'Subtasks',
                      columns = subtasks_columns)
    
    format_subtasks(workbook, writer.sheets['Subtasks'],
                    subtasks,
                    format_all,
                    header_fmt,
                    effort_fmt,
                    format_impl,
                    format_done)

    writer.save()

    return

def TMSreport(
    filename,
    jql,
    project,
    path = cmn.DEFAULT_PATH + 'Documents\\',
    WBSpath = '',
    timesheet = None):

    jira = jc.tms_login()

    print('Login OK')

    WorkItems = jc.get_WorkItems(jira, jql)

    print('Loaded ' + str(len(WorkItems)) + ' Work items')

    subtasks = jc.get_subtasks(jira, WorkItems)

    print('Loaded ' + str(len(subtasks)) + ' subtasks')

    subtasks = subtasks.sort_values('Sprint').sort_index()
    WorkItems = WorkItems.sort_index(level=0).sort_values('Sprint')
    WorkItems_enriched = WorkItems.join(WorkItems.apply(lambda row: calc_subtasks(row.name[1],WorkItems,subtasks),axis=1))

    print('Calculations OK')

    WorkItems_columns = ['Summary',
                        'Status',
                        'Priority',
                        'Assignee',
                        #'Sprint',
                        'Components',
                        'Σ Original Estimate',
                        'total_estimate',
                        'Σ Time Spent',
                        'total_remaining',
                        'total_fcast_plan_rate',
                        'total_spent_rate',
                        'total_rate',
                        'nonqa_estimate',
                        'nonqa_fact',
                        'nonqa_remaining',
                        'nonqa_spent_rate',
                        'nonqa_rate',
                        'qa_estimate',
                        'qa_fact',
                        'qa_remaining',
                        'qa_spent_rate',
                        'qa_rate',
                        'Due Date',
                        'Fix Version/s',
                        'Labels',
                        'Sub-Tasks',
                        'Status Mapped',
                        'cnt_dev_implemented',
                        'cnt_done',
                        'cnt_done_qa',
                        'cnt_nonqa',
                        'cnt_qa',
                        'cnt_total',
                        'done_estimate',
                        'done_qa_estimate',
                        'implemented_dev_estimate',
                        'workitem_fact',
                        'Σ Remaining Estimate']
    WorkItems_labels = ['Summary',
                        'Status',
                        'Priority',
                        'Assignee',
                        #'Sprint',
                        'Components',
                        'Total plan',
                        'Total fcast',
                        'Total fact',
                        'Remaining',
                        'Total fcast/plan',
                        'Total fact/fcast',
                        'Done, %',
                        'Impl forecast',
                        'Impl fact',
                        'Impl remaining',
                        'Impl fact/fcast',
                        'Impl, %',
                        'QA forecast',
                        'QA fact',
                        'QA remaining',
                        'QA fact/fcast',
                        'QA, %',
                        'Due Date',
                        'Fix Version/s',
                        'Labels',
                        'Sub-Tasks',
                        'Status Mapped',
                        'cnt_dev_implemented',
                        'cnt_done',
                        'cnt_done_qa',
                        'cnt_nonqa',
                        'cnt_qa',
                        'cnt_total',
                        'done_estimate',
                        'done_qa_estimate',
                        'implemented_dev_estimate',
                        'workitem_fact',
                        'Σ Remaining Estimate']

    subtasks_columns = ['Summary',
                        'Status',
                        'Priority',
                        'Assignee',
                        #'Sprint',
                        'Components',
                        'Original Estimate',
                        'Time Spent',
                        'Remaining Estimate',
                        'Due Date',
                        'Fix Version/s',
                        'Labels',
                        'Step',
                        'Status Mapped']

    format_all = {'font_size': 9,
                    'bottom': 1,
                    'right': 1,
                    'top': 1,
                    'left': 1}
    header_fmt = {**format_all, **{'bold': True,
                                    'text_wrap': True,
                                    'align': 'center',
                                    'valign': 'vcenter'}}
    effort_fmt = {**format_all, **{'num_format': '0.0'}}
    percent_fmt = {**format_all, **{'num_format': '0%'}}
    
    format_alert = {**format_all, **{'bg_color': '#FFBDBD'}}
    format_warn = {**format_all, **{'bg_color': '#FFCC66'}}
    
    format_impl = {**format_all, **{'bg_color': '#D9FFB7'}}
    format_done = {**format_all, **{'bg_color': '#A3FFA3'}}

    format_bar = {'type': 'data_bar',
                  'min_type': 'num',
                  'max_type': 'num',
                  'min_value': 0,
                  'max_value': 1}
    format_bar_impl = {**format_bar, **{'bar_color': '#A1FF4B'}}
    format_bar_done = {**format_bar, **{'bar_color': '#43C800'}}

    write_report(path + filename,
                WorkItems_enriched,
                subtasks,
                WorkItems_columns,
                WorkItems_labels,
                subtasks_columns,
                format_all,
                header_fmt,
                effort_fmt,
                percent_fmt,
                format_alert,
                format_warn,
                format_impl,
                format_done,
                format_bar_impl,
                format_bar_done)

    print('Done!')

    return