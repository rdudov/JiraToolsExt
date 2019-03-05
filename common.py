DEFAULT_PATH = 'C:\\Users\\rudu0916\\'

FRs_columns = ['Summary',
                'Status',
                'Sprint',
                'HL Estimate, md',
                'total_estimate_md',
                'total_fact_md',
                'total_remaining_md',
                'total_fcast_plan_rate',
                'total_spent_rate',
                'total_rate',
                'Impl_LOE_md',
                'nonqa_estimate_md',
                'nonqa_fact_md',
                'nonqa_remaining_md',
                'nonqa_fcast_plan_rate',
                'nonqa_spent_rate',
                'nonqa_rate',
                'QA LOE, md',
                'qa_estimate_md',
                'qa_fact_md',
                'qa_remaining_md',
                'qa_fcast_plan_rate',
                'qa_spent_rate',
                'qa_rate',
                'Teams',
                'Customers',
                'Due Date',
                'Labels',
                'Fix Version/s',
                'Analysis LOE, md',
                'Bug Fixing LOE, md',
                'Build LOE, md',
                'Design LOE, md',
                'cnt_stories',
                'cnt_nonqa',
                'cnt_qa',
                'cnt_total',
                'cnt_dev_implemented',
                'cnt_done_qa',
                'cnt_done',
                'total_estimate',
                'Σ Time Spent',
                'nonqa_estimate',
                'implemented_dev_estimate',
                'done_estimate',
                'done_qa_estimate',
                'qa_estimate',
                'Status Mapped']
FRs_labels = ['Summary',
                'Status',
                'Sprint',
                'Tech LOE, md',
                'Total fcast, md',
                'Total fact, md',
                'Total remaining, md',
                'Total fcast/plan',
                'Total fact/fcast',
                'Done, %',
                'Impl LOE, md',
                'Impl forecast, md',
                'Impl fact, md',
                'Impl remaining, md',
                'Impl fcast/plan',
                'Impl fact/fcast',
                'Impl, %',
                'QA LOE, md',
                'QA forecast, md',
                'QA fact, md',
                'QA remaining, md',
                'QA fcast/plan',
                'QA fact/fcast',
                'QA, %',
                'Teams',
                'Customers',
                'Due Date',
                'Labels',
                'Fix Version/s',
                'Analysis LOE, md',
                'Bug Fixing LOE, md',
                'Build LOE, md',
                'Design LOE, md',
                'cnt_stories',
                'cnt_nonqa',
                'cnt_qa',
                'cnt_total',
                'cnt_dev_implemented',
                'cnt_done_qa',
                'cnt_done',
                'total_estimate',
                'Σ Time Spent',
                'nonqa_estimate',
                'implemented_dev_estimate',
                'done_estimate',
                'done_qa_estimate',
                'qa_estimate',
                'Status Mapped']

Stories_columns = ['Summary',
                    'Status',
                    'Team',
                    'Assignee',
                    'Sprint',
                    'total_plan',
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
                    'story_fact',
                    'Σ Remaining Estimate']
Stories_labels = ['Summary',
                    'Status',
                    'Team',
                    'Assignee',
                    'Sprint',
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
                    'story_fact',
                    'Σ Remaining Estimate']

subtasks_columns = ['Summary',
                    'Status',
                    'Assignee',
                    'Sprint',
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

format_inprogress = {**format_all, **{'bg_color': '#FFF9C1'}}
format_impl = {**format_all, **{'bg_color': '#AAFF5D'}} #D9FFB7
format_done = {**format_all, **{'bg_color': '#A3FFA3'}}
format_cancelled = {**format_all, **{'bg_color': '#D9D9D9'}}

format_bar = {'type': 'data_bar',
                'min_type': 'num',
                'max_type': 'num',
                'min_value': 0,
                'max_value': 1}
format_bar_impl = {**format_bar, **{'bar_color': '#A1FF4B'}}
format_bar_done = {**format_bar, **{'bar_color': '#43C800'}}
format_bar_blue = {**format_bar, **{'bar_color': '#638EC6'}}
format_bar_green = {**format_bar, **{'bar_color': '#63C384'}}

#99 195 132
#99 142 198
    
format_addcope = {**format_all, **{'bg_color': '#A3FFA3'}}
format_descope = {**format_all, **{'bg_color': '#A3FFA3'}}

format_good = {**format_all, **{'bg_color': '#A3FFA3'}}
format_average = {**format_all, **{'bg_color': '#FFF9C1'}}
format_bad = {**format_all, **{'bg_color': '#FFBDBD'}}


def nvl(i,o):
    if i is None:
        return o
    else:
       return i

def pretty(d, indent=0):
   if isinstance(d, dict):
       for key, value in d.items():
          print('\t' * indent + str(key))
          if isinstance(value, dict) | isinstance(value, list) :
             pretty(value, indent+1)
          else:
             print('\t' * (indent+1) + str(value))
   if isinstance(d, list):
       for item in d:
          if isinstance(item, dict) | isinstance(item, list) :
             pretty(item, indent+1)
          else:
             print('\t' * (indent+1) + str(item))

def pretty1(d, indent=0, parent='', exclude=set()):
   if isinstance(d, dict):
       for key, value in d.items():
          if parent + '_' + key not in exclude:
            print('\t' * indent + str(key))
            exclude.add(parent + '_' + key)
          if isinstance(value, dict) | isinstance(value, list) :
             pretty1(value, indent=indent+1, parent=key, exclude=exclude)
   if isinstance(d, list):
       for item in d:
          if isinstance(item, dict) | isinstance(item, list) :
             pretty1(item, indent=indent+1, parent=parent, exclude=exclude)

def pretty2(d, parent=[]):
   res = set()
   if isinstance(d, dict):
       for key, value in d.items():
          res.add('/'.join(parent+[key]))
          if isinstance(value, dict) | isinstance(value, list) :
             res = res.union(pretty2(value, parent=parent+[key]))
   if isinstance(d, list):
       for item in d:
          if isinstance(item, dict) | isinstance(item, list) :
             res = res.union(pretty2(item, parent=parent))
   return res
 
