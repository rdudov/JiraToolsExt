import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell
import pandas.io.formats.excel
import jiraConnector as jc
import WBStools as wbs
import TERtools as ter
import numpy as np
import re
import common as cmn

def original_estimate(row):
    return max(row['Original Estimate'], row['Time Spent'] + row['Remaining Estimate'])

def remaining_estimate(row):
    if (row['Status Mapped']=='Done') | (row['Status Mapped']=='In QA'):
        return 0
    else:
       return row['Remaining Estimate']

def calc_subtasks(Story, Stories, subtasks):
    
    #print('calc_subtasks in Story=' + Story)

    Story_row = Stories[Stories.index.get_level_values(1) == Story].iloc[0]
    
    story_estimate = original_estimate(Story_row)
    story_fact = Story_row['Time Spent']
    total_fact = Story_row['Σ Time Spent']
    
    story_done_cnt = 0
    story_done_estimate = 0
    story_implemented_cnt = 0
    story_implemented_estimate = 0
    story_remaining_estimate = 0

    subtasks_c = subtasks[(subtasks.index.get_level_values(1) == Story) & (subtasks['Status Mapped']!='Cancelled')]
    subtasks_c_all = subtasks[subtasks.index.get_level_values(1) == Story]

    cnt_subtasks = len(subtasks_c)
    cnt_subtasks_all = len(subtasks_c_all)

    if story_estimate > 0 or cnt_subtasks == 0:
        story_cnt = 1
        if Story_row['Status Mapped'] == 'Done':
            story_done_cnt = 1
            story_implemented_cnt = 1
            story_implemented_estimate = story_estimate
            story_done_estimate = story_estimate
        elif Story_row['Status Mapped'] == 'In QA':
            story_implemented_cnt = 1
            story_implemented_estimate = story_estimate
        elif Story_row['Status Mapped'] == 'Cancelled':
            story_estimate = 0
        else:
            story_remaining_estimate = Story_row['Remaining Estimate']
    else:
        story_cnt = 0

    cnt_total = cnt_subtasks + story_cnt

    cnt_nonqa_subtasks = len(subtasks_c[subtasks_c['Step']!='QA'])
    cnt_nonqa = cnt_nonqa_subtasks + story_cnt

    cnt_qa = cnt_total-cnt_nonqa

    cnt_done_subtasks = len(subtasks_c[subtasks_c['Status Mapped']=='Done'])
    cnt_done = cnt_done_subtasks + story_done_cnt

    cnt_done_qa = len(subtasks_c[(subtasks_c['Step']=='QA') & (subtasks_c['Status Mapped']=='Done')])

    cnt_dev_implemented_subtasks = len(subtasks_c[(subtasks_c['Step']!='QA') & subtasks_c['Status Mapped'].isin(['In QA', 'Done'])])
    cnt_dev_implemented = cnt_dev_implemented_subtasks + story_implemented_cnt

    if cmn.nvl(Story_row['WBS Dev LOE'],0) > 0:
        total_plan = Story_row['WBS Dev LOE']
    else:
        total_plan = Story_row['Σ Original Estimate']

    if cnt_subtasks > 0:
        total_estimate = sum(subtasks_c.apply(original_estimate,axis=1)) + story_estimate
        total_remaining = sum(subtasks_c.apply(remaining_estimate,axis=1)) + story_remaining_estimate
    elif cnt_subtasks_all > 0:
        total_estimate = sum(subtasks_c_all.apply(original_estimate,axis=1)) + story_estimate
        total_remaining = story_remaining_estimate
    else:
        total_estimate = story_estimate
        total_remaining = story_remaining_estimate

    total_estimate_for_rate = total_estimate

    if total_plan > 0:
        if total_estimate/total_plan <= 0.2:
            total_estimate_for_rate = total_plan

    if cnt_nonqa_subtasks > 0:
        nonqa_estimate = sum(subtasks_c[subtasks_c['Step']!='QA'].apply(original_estimate,axis=1)) + story_estimate
        nonqa_fact = sum(subtasks_c[subtasks_c['Step']!='QA'].apply(lambda row: row['Time Spent'],axis=1)) + story_fact
        nonqa_remaining = sum(subtasks_c[subtasks_c['Step']!='QA'].apply(remaining_estimate,axis=1)) + story_remaining_estimate
    else:
        nonqa_estimate = story_estimate
        nonqa_fact = story_fact
        nonqa_remaining = story_remaining_estimate

    qa_estimate = total_estimate - nonqa_estimate
    qa_fact = total_fact - nonqa_fact
    qa_remaining = total_remaining - nonqa_remaining

    if cnt_done_subtasks > 0:
        done_estimate = sum(subtasks_c[subtasks_c['Status Mapped']=='Done'].apply(original_estimate,axis=1)) + story_done_estimate
    else:
        done_estimate = story_done_estimate

    if cnt_done_qa > 0:
        done_qa_estimate = sum(subtasks_c[(subtasks_c['Step']=='QA') & (subtasks_c['Status Mapped']=='Done')].apply(original_estimate,axis=1))
    else:
        done_qa_estimate = 0

    if cnt_dev_implemented_subtasks > 0:
        implemented_dev_estimate = sum(subtasks_c[(subtasks_c['Step']!='QA') & subtasks_c['Status Mapped'].isin(['In QA', 'Done'])].apply(original_estimate,axis=1)) + story_implemented_estimate
    else:
        implemented_dev_estimate = story_implemented_estimate

    if total_estimate_for_rate > 0:
        total_rate = done_estimate / total_estimate_for_rate
        total_spent_rate = total_fact / total_estimate_for_rate
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
    
    if total_plan > 0:
        total_fcast_plan_rate = total_estimate / total_plan
    else:
        total_fcast_plan_rate = 0
    
    if cnt_total == 0 or total_estimate == 0:
        if Story_row['Status Mapped'] == 'Done':
            total_rate = 1
            nonqa_rate = 1
            qa_rate = 1
        elif Story_row['Status Mapped'] == 'In QA':
            total_rate = 0
            nonqa_rate = 1
            qa_rate = 0

    return pd.Series({'cnt_total': cnt_total,
                     'cnt_nonqa': cnt_nonqa,
                     'cnt_qa': cnt_qa,
                     'cnt_done': cnt_done,
                     'cnt_done_qa': cnt_done_qa,
                     'cnt_dev_implemented': cnt_dev_implemented,
                     'total_plan' : total_plan,
                     'total_estimate': total_estimate,
                     'nonqa_estimate': nonqa_estimate,
                     'qa_estimate': qa_estimate,
                     'total_fcast_plan_rate': total_fcast_plan_rate,
                     'story_fact': story_fact,
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

def FRs_rate(FR, Stories):

    cnt_stories = len(Stories[Stories.index.get_level_values(0).isin([FR.name])])

    Impl_LOE_md = FR['Design LOE, md'] + FR['Build LOE, md']
    total_fact_md = FR['Σ Time Spent']/8
    nonqa_fact_md = FR['nonqa_fact']/8
    qa_fact_md = FR['qa_fact']/8
    nonqa_estimate_md=0
    qa_estimate_md=0
    total_estimate_md=0
    nonqa_remaining_md=0
    qa_remaining_md=0
    total_remaining_md=0
    total_fcast_plan_rate=0
    nonqa_fcast_plan_rate=0
    qa_fcast_plan_rate=0
    total_spent_rate=0
    nonqa_spent_rate=0
    qa_spent_rate=0
    total_rate = 0
    nonqa_rate = 0
    qa_rate = 0

    if FR['cnt_total'] == 0:
        if FR['Status Mapped'] == 'Done':
            total_rate = 1
            nonqa_rate = 1
            qa_rate = 1
        elif FR['Status Mapped'] == 'In QA':
            total_rate = 0
            nonqa_rate = 1
            qa_rate = 0

    elif FR['total_estimate'] == 0:
        if FR['cnt_nonqa'] > 0:
            nonqa_rate = FR['cnt_dev_implemented'] / FR['cnt_nonqa']

        if FR['cnt_qa'] > 0:
            qa_rate = FR['cnt_done_qa'] / FR['cnt_qa']

        total_rate = FR['cnt_done'] / FR['cnt_total']

    else:
        nonqa_estimate_md = FR['nonqa_estimate']/8
        qa_estimate_md = FR['qa_estimate']/8
        total_estimate_md = FR['total_estimate']/8
        nonqa_remaining_md = FR['nonqa_remaining']/8
        qa_remaining_md = FR['qa_remaining']/8
        total_remaining_md = FR['total_remaining']/8

        total_rate = FR['done_estimate'] / FR['total_estimate']
        total_spent_rate = FR['Σ Time Spent'] / FR['total_estimate']

        if FR['HL Estimate, md'] > 0:
            total_fcast_plan_rate = total_estimate_md / FR['HL Estimate, md']
        else:
            total_fcast_plan_rate = 0

        if Impl_LOE_md > 0:
            nonqa_fcast_plan_rate = nonqa_estimate_md / Impl_LOE_md
        else:
            nonqa_fcast_plan_rate = 0

        if FR['QA LOE, md'] > 0:
            qa_fcast_plan_rate = qa_estimate_md / FR['QA LOE, md']
        else:
            qa_fcast_plan_rate = 0

        if FR['nonqa_estimate'] > 0:
            nonqa_rate = FR['implemented_dev_estimate'] / FR['nonqa_estimate']
            nonqa_spent_rate = FR['nonqa_fact'] / FR['nonqa_estimate']
        else:
            nonqa_spent_rate = 0

        if FR['qa_estimate'] > 0:
            qa_rate = FR['done_qa_estimate'] / FR['qa_estimate']
            qa_spent_rate = FR['qa_fact'] / FR['qa_estimate']
        else:
            qa_spent_rate = 0

    return pd.Series({'cnt_stories': cnt_stories,
                    'Impl_LOE_md': Impl_LOE_md,
                    'nonqa_estimate_md': nonqa_estimate_md,
                    'qa_estimate_md': qa_estimate_md,
                    'total_estimate_md': total_estimate_md,
                    'nonqa_remaining_md': nonqa_remaining_md,
                    'qa_remaining_md': qa_remaining_md,
                    'total_remaining_md': total_remaining_md,
                    'total_fcast_plan_rate': total_fcast_plan_rate,
                    'nonqa_fcast_plan_rate': nonqa_fcast_plan_rate,
                    'qa_fcast_plan_rate': qa_fcast_plan_rate,
                    'total_fact_md': total_fact_md,
                    'nonqa_fact_md': nonqa_fact_md,
                    'qa_fact_md': qa_fact_md,
                    'total_spent_rate': total_spent_rate,
                    'nonqa_spent_rate': nonqa_spent_rate,
                    'qa_spent_rate': qa_spent_rate,
                    'total_rate': total_rate,
                    'nonqa_rate': nonqa_rate,
                    'qa_rate': qa_rate})

def format_FRs(workbook,
               worksheet,
               number_rows,
               project,
               timesheet):

    #http://xlsxwriter.readthedocs.io/format.html

    format_all = workbook.add_format(cmn.format_all)
    header_fmt = workbook.add_format(cmn.header_fmt)
    effort_fmt = workbook.add_format(cmn.effort_fmt)
    percent_fmt = workbook.add_format(cmn.percent_fmt)

    worksheet.set_column('A:A', 11, format_all)
    worksheet.set_column('B:B', 40, format_all)
    worksheet.set_column('C:C', 11, format_all)
    worksheet.set_column('D:D', 18, format_all)
    worksheet.set_column('E:H', 8, effort_fmt)
    worksheet.set_column('I:K', 8, percent_fmt)
    worksheet.set_column('L:O', 8, effort_fmt)
    worksheet.set_column('P:R', 8, percent_fmt)
    worksheet.set_column('S:V', 8, effort_fmt)
    worksheet.set_column('W:Y', 8, percent_fmt)
    worksheet.set_column('Z:AA', 13, format_all)
    worksheet.set_column('AB:BC', 8, format_all)

    worksheet.set_row(0, None, header_fmt)

    format_alert = workbook.add_format(cmn.format_alert)
    format_warn = workbook.add_format(cmn.format_warn)
    
    format_impl = workbook.add_format(cmn.format_impl)
    format_done = workbook.add_format(cmn.format_done)
    format_cancelled = workbook.add_format(cmn.format_cancelled)
    
    #Fact exceeds plan (forecast rate * spent rate) for over 20%
    worksheet.conditional_format('G2:G{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=($I2*$J2)>=1.2', 'format': format_alert})
    worksheet.conditional_format('N2:N{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=($P2*$Q2)>=1.2', 'format': format_alert})
    worksheet.conditional_format('U2:U{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=($W2*$X2)>=1.2', 'format': format_alert})
    
    #Cancelled
    worksheet.conditional_format('A2:BC{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=$AW2="Cancelled"', 'format': format_cancelled})
    
    #Feature is completed
    worksheet.conditional_format('A2:BC{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=$K2=1', 'format': format_done})
    
    #Development is completed
    worksheet.conditional_format('A2:BC{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=$R2=1', 'format': format_impl})

    #Progress pars
    worksheet.conditional_format('K2:K{}'.format(number_rows+2), cmn.format_bar_done)
    worksheet.conditional_format('R2:R{}'.format(number_rows+2), cmn.format_bar_impl)
    worksheet.conditional_format('Y2:Y{}'.format(number_rows+2), cmn.format_bar_impl)
    
    #Gained effort exceeds gained progress for over then 20%
    worksheet.conditional_format('Q2:Q{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=($Q2-$R2)>=0.2', 'format': format_warn})
    
    #Forecast exceeds plan for over than 20%
    worksheet.conditional_format('I2:I{}'.format(number_rows+2),
                                {'type': 'cell',
                                'criteria': '>=',
                                'value': 1.2,
                                'format': format_alert,
                                'multi_range': 'J2:J{} Q2:Q{} X2:X{}'.format(number_rows+2, number_rows+2, number_rows+2)})

    total_effort_fmt = workbook.add_format({**cmn.effort_fmt, **{'bold': True, 'top': 6}})
    total_percent_fmt = workbook.add_format({**cmn.percent_fmt, **{'bold': True, 'top': 6}})

    worksheet.write_string(number_rows+1, 3, "Total",total_effort_fmt)

    for column in [4, 5, 6, 7, 11, 12, 13, 14, 18, 19, 20, 21]:
        worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, column),
                                "=SUM({:s}:{:s})".format(xl_rowcol_to_cell(1, column), xl_rowcol_to_cell(number_rows, column)),
                                total_effort_fmt)
        
    #Total fcast/plan
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 8),
                            "={:s}/{:s}".format(xl_rowcol_to_cell(number_rows+1, 5),
                                                xl_rowcol_to_cell(number_rows+1, 4)),
                            total_percent_fmt)
    #Total fact/fcast
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 9),
                            "={:s}/{:s}".format(xl_rowcol_to_cell(number_rows+1, 6),
                                                xl_rowcol_to_cell(number_rows+1, 5)),
                            total_percent_fmt)
    #Done, %
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 10),
                            "=SUM({:s}:{:s})/SUM({:s}:{:s})".format(xl_rowcol_to_cell(1, 45),
                                                                    xl_rowcol_to_cell(number_rows, 45),
                                                                    xl_rowcol_to_cell(1, 41),
                                                                    xl_rowcol_to_cell(number_rows, 41)),
                            total_percent_fmt)
    #Impl fcast/plan
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 15),
                            "={:s}/{:s}".format(xl_rowcol_to_cell(number_rows+1, 12),
                                                xl_rowcol_to_cell(number_rows+1, 11)),
                            total_percent_fmt)
    #Impl fact/fcast
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 16),
                            "={:s}/{:s}".format(xl_rowcol_to_cell(number_rows+1, 13),
                                                xl_rowcol_to_cell(number_rows+1, 12)),
                            total_percent_fmt)
    #Impl, %
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 17),
                            "=SUM({:s}:{:s})/SUM({:s}:{:s})".format(xl_rowcol_to_cell(1, 44),
                                                                    xl_rowcol_to_cell(number_rows, 44),
                                                                    xl_rowcol_to_cell(1, 43),
                                                                    xl_rowcol_to_cell(number_rows, 43)),
                            total_percent_fmt)
    #QA fcast/plan
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 22),
                            "={:s}/{:s}".format(xl_rowcol_to_cell(number_rows+1, 19),
                                                xl_rowcol_to_cell(number_rows+1, 18)),
                            total_percent_fmt)
    #QA fact/fcast
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 23),
                            "={:s}/{:s}".format(xl_rowcol_to_cell(number_rows+1, 20),
                                                xl_rowcol_to_cell(number_rows+1, 19)),
                            total_percent_fmt)
    #QA, %
    worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, 24),
                            "=SUM({:s}:{:s})/SUM({:s}:{:s})".format(xl_rowcol_to_cell(1, 46),
                                                                    xl_rowcol_to_cell(number_rows, 46),
                                                                    xl_rowcol_to_cell(1, 47),
                                                                    xl_rowcol_to_cell(number_rows, 47)),
                            total_percent_fmt)
    
    if timesheet is not None:
        worksheet.write_string(number_rows+2, 3, "Total ERP",total_effort_fmt)
        worksheet.write_number(number_rows+2, 4, ter.PROJECTS_INFO[project]['projectLOE'], total_effort_fmt)
        worksheet.write_number(number_rows+2, 6, timesheet['MD'].sum(), total_effort_fmt)
    
    return

def format_Stories(workbook,
                   worksheet,
                   FRs,
                   Stories,
                   number_columns,
                   project,
                   timesheet):
    
    number_rows = len(Stories.index)

    format_all = workbook.add_format(cmn.format_all)
    header_fmt = workbook.add_format(cmn.header_fmt)
    effort_fmt = workbook.add_format(cmn.effort_fmt)
    percent_fmt = workbook.add_format(cmn.percent_fmt)

    worksheet.set_column('A:B', 11, format_all)
    worksheet.set_column('C:C', 40, format_all)
    worksheet.set_column('D:D', 11, format_all)
    worksheet.set_column('E:F', 15, format_all)
    worksheet.set_column('G:G', 18, format_all)
    worksheet.set_column('H:K', 8, effort_fmt)
    worksheet.set_column('L:N', 8, percent_fmt)
    worksheet.set_column('O:Q', 8, effort_fmt)
    worksheet.set_column('R:S', 8, percent_fmt)
    worksheet.set_column('T:V', 8, effort_fmt)
    worksheet.set_column('W:X', 8, percent_fmt)
    worksheet.set_column('Y:AN',8, format_all)

    worksheet.set_row(0, None, header_fmt)
    
    format_alert = workbook.add_format(cmn.format_alert)
    format_warn = workbook.add_format(cmn.format_warn)
    
    format_impl = workbook.add_format(cmn.format_impl)
    format_done = workbook.add_format(cmn.format_done)
    format_cancelled = workbook.add_format(cmn.format_cancelled)
    
    #Fact exceeds plan (forecast rate * spent rate) for over 20%
    worksheet.conditional_format('J2:J{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=($L2*$M2)>=1.2', 'format': format_alert})

    #Cancelled
    worksheet.conditional_format('A2:AN{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=$AC2="Cancelled"', 'format': format_cancelled})

    #Story is completed
    worksheet.conditional_format('A2:AN{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=$N2=1', 'format': format_done})
    
    #Development is completed
    worksheet.conditional_format('A2:AN{}'.format(number_rows+1), {'type': 'formula', 'criteria': '=$S2=1', 'format': format_impl})
    
    #Progress pars
    worksheet.conditional_format('N2:N{}'.format(number_rows+2), cmn.format_bar_done)
    worksheet.conditional_format('S2:S{}'.format(number_rows+2), cmn.format_bar_impl)
    worksheet.conditional_format('X2:X{}'.format(number_rows+2), cmn.format_bar_impl)
    
    #Gained effort exceeds gained progress for over then 20%
    worksheet.conditional_format('R2:R{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=($R2-$S2)>=0.2', 'format': format_warn})
    
    #Forecast exceeds plan for over than 20%
    worksheet.conditional_format('L2:L{}'.format(number_rows+2),
                                {'type': 'cell',
                                'criteria': '>=',
                                'value': 1.2,
                                'format': format_alert})

    total_effort_fmt = workbook.add_format({**cmn.effort_fmt, **{'bold': True, 'top': 6}})
    total_percent_fmt = workbook.add_format({**cmn.percent_fmt, **{'bold': True, 'top': 6}})

    worksheet.write_string(number_rows+1, 6, "Total",total_effort_fmt)

    for column in [7, 8, 9, 10, 14, 15, 16, 19, 20, 21]:
        worksheet.write_formula(xl_rowcol_to_cell(number_rows+1, column),
                                "=SUM({:s}:{:s})".format(xl_rowcol_to_cell(1, column), xl_rowcol_to_cell(number_rows, column)),
                                total_effort_fmt)
        
    if FRs is not None:
        for row in range(number_rows):
            if cmn.nvl(Stories.iloc[row].name[0],'') != '':
                try:
                    comment = (FRs.loc[Stories.iloc[row].name[0]]['Summary'] + '\n' +
                               str(int(FRs.loc[Stories.iloc[row].name[0]]['nonqa_rate'] * 100)) + '% - ' +
                               str(round(FRs.loc[Stories.iloc[row].name[0]]['HL Estimate, md'],1)) + '/' +
                               str(round(FRs.loc[Stories.iloc[row].name[0]]['total_fact_md'],1)) + '/' +
                               str(round(FRs.loc[Stories.iloc[row].name[0]]['total_remaining_md'],1)))
                    worksheet.write_comment('A{}'.format(row+2),comment)
                except:
                    print('Cannot create comment for Story ' + Stories.iloc[row].name[1])

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
        
    if timesheet is not None:
        worksheet.write_string(number_rows+2, 6, "Total ERP",total_effort_fmt)
        worksheet.write_number(number_rows+2, 7, ter.PROJECTS_INFO[project]['projectLOE'], total_effort_fmt)
        worksheet.write_number(number_rows+2, 9, timesheet['MD'].sum(), total_effort_fmt)

    worksheet.autofilter(0,0,number_rows+1,number_columns+1)

    return

def format_subtasks(workbook,
                   worksheet,
                   number_rows,
                   number_columns):
    
    format_all = workbook.add_format(cmn.format_all)
    header_fmt = workbook.add_format(cmn.header_fmt)
    effort_fmt = workbook.add_format(cmn.effort_fmt)

    worksheet.set_column('A:C', 11, format_all)
    worksheet.set_column('D:D', 40, format_all)
    worksheet.set_column('E:E', 11, format_all)
    worksheet.set_column('F:F', 15, format_all)
    worksheet.set_column('G:G', 18, format_all)
    worksheet.set_column('H:J', 8, effort_fmt)
    worksheet.set_column('K:O',8, format_all)

    worksheet.set_row(0, None, header_fmt)
    
    format_impl = workbook.add_format(cmn.format_impl)
    format_done = workbook.add_format(cmn.format_done)
    format_cancelled = workbook.add_format(cmn.format_cancelled)
    
    #Resolved
    worksheet.conditional_format('A2:O{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=$O2="In QA"', 'format': format_impl})

    #Closed
    worksheet.conditional_format('A2:O{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=$O2="Done"', 'format': format_done})
    
    #Cancelled
    worksheet.conditional_format('A2:O{}'.format(number_rows+2), {'type': 'formula', 'criteria': '=$O2="Cancelled"', 'format': format_cancelled})

    worksheet.autofilter(0,0,number_rows+1,number_columns+2)

    return

def write_report(path,
                 FRs,
                 Stories,
                 subtasks,
                 project,
                 FRs_columns,
                 FRs_labels,
                 timesheet):
    
    writer = pd.ExcelWriter(path, engine='xlsxwriter')

    pd.io.formats.excel.header_style = None

    FRs.to_excel(writer,
                index_label = ['Key'],
                merge_cells=False,
                freeze_panes=(1,2),
                sheet_name = 'FRs',
                columns = FRs_columns,
                header = FRs_labels)

    workbook = writer.book
    
    format_FRs(workbook, writer.sheets['FRs'],
                len(FRs.index),
                project,
                timesheet)

    Stories.to_excel(writer,
                    index_label = ['FR','Key'],
                    merge_cells=False,
                    freeze_panes=(1,2),
                    sheet_name = 'Stories',
                    columns = cmn.Stories_columns,
                    header = cmn.Stories_labels)
    
    format_Stories(workbook, writer.sheets['Stories'],
                   FRs,
                   Stories,
                   len(cmn.Stories_columns),
                   project,
                   None)

    subtasks.to_excel(writer,
                      index_label = ['FR','Story','Key'],
                      merge_cells=False,
                      freeze_panes=(1,3),
                      sheet_name = 'Subtasks',
                      columns = cmn.subtasks_columns)
    
    format_subtasks(workbook, writer.sheets['Subtasks'],
                    len(subtasks.index),
                    len(cmn.subtasks_columns))

    writer.save()

    return


def write_report_short(path,
                     Stories,
                     subtasks,
                     project,
                     timesheet = None):
    
    writer = pd.ExcelWriter(path, engine='xlsxwriter')

    pd.io.formats.excel.header_style = None

    workbook = writer.book

    Stories.to_excel(writer,
                    index_label = ['FR','Key'],
                    merge_cells=False,
                    freeze_panes=(1,2),
                    sheet_name = 'Stories',
                    columns = cmn.Stories_columns,
                    header = cmn.Stories_labels)
    
    format_Stories(workbook, writer.sheets['Stories'],
                   None,
                   Stories,
                   len(cmn.Stories_columns),
                   project,
                   timesheet = timesheet)

    subtasks.to_excel(writer,
                      index_label = ['FR','Story','Key'],
                      merge_cells=False,
                      freeze_panes=(1,3),
                      sheet_name = 'Subtasks',
                      columns = cmn.subtasks_columns)
    
    format_subtasks(workbook, writer.sheets['Subtasks'],
                    len(subtasks.index),
                    len(cmn.subtasks_columns))

    writer.save()

    return

def FRreport(
    filename,
    FR_filter,
    Stories_filter,
    project,
    path = cmn.DEFAULT_PATH + 'Documents\\',
    add_clause = '',
    WBSpath = '',
    WBStag = 'WBS',
    timesheet = None):

    jira = jc.psup_login()

    print('Login OK')

    if len(FR_filter) > 0:
        FRs = jc.get_FRs(jira, FR_filter, WBStag=WBStag)

        print('Loaded ' + str(len(FRs)) + ' FRs')

        Stories = jc.get_Stories(jira, FRs=FRs.index, add_clause=add_clause, WBStag=WBStag)
    else:
        Stories = jc.get_Stories(jira, filter_id=Stories_filter, WBStag=WBStag)

    print('Loaded ' + str(len(Stories)) + ' Stories')

    subtasks = jc.get_subtasks(jira, Stories)

    print('Loaded ' + str(len(subtasks)) + ' subtasks')

    subtasks = subtasks.sort_values('Sprint').sort_index()
    Stories = Stories.sort_index(level=0).sort_values('Sprint')
    Stories_enriched = Stories.join(Stories.apply(lambda row: calc_subtasks(row.name[1],Stories, subtasks),axis=1))

    if len(FR_filter) > 0:
        FRs = FRs.sort_values('Sprint')
        FRs_enriched = FRs.join(Stories_enriched[[
            'cnt_dev_implemented',
            'cnt_done',
            'cnt_done_qa',
            'cnt_nonqa',
            'cnt_qa',
            'cnt_total',
            'Σ Time Spent',
            'nonqa_fact',
            'qa_fact',
            'nonqa_remaining',
            'qa_remaining',
            'total_remaining',
            'done_estimate',
            'done_qa_estimate',
            'implemented_dev_estimate',
            'nonqa_estimate',
            'qa_estimate',
            'total_estimate']].groupby(level=0).sum()).fillna(0)
        FRs_enriched = FRs_enriched.join(FRs_enriched.apply(lambda row: FRs_rate(row,Stories),axis=1))

        WBS_columns = ['Module',
                       'Sprint wbs',
                       'Customer wbs',
                       'Tech LOE wbs',
                       'Impl LOE wbs',
                       'QA LOE wbs']
        names = ['Scope of Work',
                'Customer wbs',
                'Module',
                'Sprint wbs',
                'Estimated By',
                'Details/Assumptions',
                'Analysis LOE wbs',
                'Bug Fixing LOE wbs',
                'Design LOE wbs', 
                'Build LOE wbs',
                'QA LOE wbs',
                'Tech LOE wbs']
        subset = ['Customer wbs',
                'Module',
                'Sprint wbs',
                'Estimated By']

        if WBSpath != '':
            WBS = wbs.loadWBS(WBSpath,names,subset)
            FRs_enriched = FRs_enriched.join(WBS[WBS_columns])
        else:
            FRs_enriched = FRs_enriched.reindex(columns = np.append(FRs_enriched.columns.values,WBS_columns))

    print('Calculations OK')

    if len(FR_filter) > 0:
        write_report(path + filename,
                 FRs_enriched,
                 Stories_enriched,
                 subtasks,
                 project,
                 np.append(cmn.FRs_columns,WBS_columns),
                 np.append(cmn.FRs_labels,WBS_columns),
                 timesheet)
    else:
        write_report_short(path + filename,
                             Stories_enriched,
                             subtasks,
                             project,
                             timesheet)

    print('Done!')

    return