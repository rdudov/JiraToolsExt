import FRreport as FR
import TMSreport as TMS
import datetime
import TERtools as ter
import common as cmn

def runReport(project):
    
    print('Run report for ' + project)

    type = 'PSUP'
    filename = 'report_' + project + '_' + datetime.datetime.now().strftime("%Y-%m-%d") + '.xlsx'
    
    Stories_filter = ''
    add_clause = ''
    WBSpath = ''
    timesheet = None

    if project in ter.PROJECTS_SHORTLIST:
        timesheet = ter.getTimesheetForProject(project)        
        print('Timesheets are loaded')

    if project == 'ES20_R3':
        FR_filter = '57567'
        
    elif project == 'VDOCR4':
        FR_filter = '56536'
        add_clause = 'fixVersion = "Release 4"'

    elif project == 'VDOCR5':
        FR_filter = '59394'
        add_clause = 'fixVersion = "Release 5"'
        
    elif project == 'ESO91':
        FR_filter = ''
        Stories_filter = '58846'

    elif project == 'CM91':
        FR_filter = ''
        Stories_filter = '58848'

    elif project == 'CM911':
        FR_filter = ''
        Stories_filter = '58848'

    elif project == 'OPTDPNV':
        jql = 'project = "NC.ENG.Optus DP Naming Versioning" and type = "Work Item"'
        type = 'TMS'

    #OLD
    else:
        if project == 'UUI_80':
            FR_filter = '52191'

        elif project == 'UUI_81':
            FR_filter = '55044'
            add_clause = 'project in (MANOGUI, UMBRUI, CHOM)'
            WBSpath = cmn.DEFAULT_PATH + 'Desktop\\SDN_NFV\\HOM\\UI\\UI Proposals R 8.1\\NetCracker_R8.1_MANO_Umbrella_UI_WBS_v4.1.xlsx'

        elif project == 'VDOC_R3':
            FR_filter = ''
            Stories_filter = '54956'

        elif project == 'TFA_81':
            FR_filter = '54073'
            add_clause = '(fixVersion="MANO 8.1" or fixVersion is empty)'

        elif project == 'TFA_80':
            FR_filter = '51495'
        
        elif project == 'IPAM_81':
            FR_filter = '54427'
            
        elif project == 'ES20_R2':
            FR_filter = '57569'

        elif project == 'ES20_R1':
            FR_filter = '57570'
   
        elif project == 'UUI_90':
            FR_filter = '56493'
            add_clause = 'project in (MANOGUI, UMBRUI)'
            #WBSpath = 'C:\\Users\\rudu0916\\Desktop\\SDN_NFV\\HOM\\UI\\UI Proposals R 9.0\\NetCracker_R9.0_MANO_Umbrella_UI_WBS_v9.0_v10_rs.xlsx'
    
        elif project == 'TFA_90':
            FR_filter = '56924'

        elif project == 'IPAM_90':
            FR_filter = '56429'

        elif project == 'DT_Intraselect':
            FR_filter = ''
            Stories_filter = '56232'

    if type == 'PSUP':
        FR.FRreport(
            filename,
            FR_filter,
            Stories_filter,
            project,
            add_clause = add_clause,
            WBSpath = WBSpath,
            WBStag = project + '_WBS',
            timesheet = timesheet)
    elif type == 'TMS':
        TMS.TMSreport(
            filename,
            jql,
            project,
            WBSpath = WBSpath,
            timesheet = timesheet)



#runReport('UUI_81')
#runReport('UUI_80')
#runReport('TFA_81')
#runReport('TFA_80')
#runReport('IPAM_81')
#runReport('VDOC_R3')
#runReport('ES20_R2')
#runReport('ES20_R1')
#runReport('UUI_90')
#runReport('IPAM_90')
#runReport('TFA_90')
#runReport('DT_Intraselect')
#runReport('OPTDPNV')


#runReport('VDOCR4')
runReport('VDOCR5')
#runReport('ES20_R3')
#runReport('ESO91')
#runReport('CM91')
#runReport('CM911')