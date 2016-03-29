__author__ = 'aviran-k'

from win32com.client import Dispatch
import xlsxwriter
import openpyxl
from tkFileDialog import askdirectory
from Tkinter import Tk
import os
from datetime import time
from datetime import date
import csv
import re

###### FUNCTIONS ######


# Convert a float to percentage, for usage in cases like boot rates
# @param frac - the float
# @param digits_after_decimal_point - the desired amount of digits after decimal point
# @return - the float as percentage
def convert_fraction_to_percentage(frac, digits_after_decimal_point=None):
    percentage = frac*100
    if digits_after_decimal_point is None:
        return str(percentage)+'%'
    else:
        return str(round(percentage, digits_after_decimal_point))+'%'

# Extract HTML table from a CSV file
# @param csv_file - the CSV file to create the table from
# @return - the HTML table
def csv_to_html_table(csv_file, has_header=False, convert_fracs_to_percentage=False, digs_after_dec_point=None):
    global exceptions
    html_table = "<table>"
    cell_tags = {'open': '<td>', 'close': '</td>'}
    with open(csv_file, 'r') as handle:
        contents = csv.reader(handle)
        for row in contents:
            html_table += '<tr>'
            if has_header:
                cell_tags = {'open': '<th>', 'close': '</th>'}
            for i in row:
                if convert_fracs_to_percentage:
                    try:
                        cell_str = convert_fraction_to_percentage(float(i), digits_after_decimal_point=digs_after_dec_point)
                    except:
                        cell_str = i
                else:
                    cell_str = i
                html_table += cell_tags['open']
                html_table += cell_str
                html_table += cell_tags['close']
            html_table += '</tr>'
            if has_header:
                cell_tags = {'open': '<td>', 'close': '</td>'}
                has_header = False
    html_table += '</table>'
    return html_table


# Parse nr_info.txt into dictionary
# @param dir - directory of nr_info.txt (Nightrun directory)
# @return - dictionary that contains data and configurations of nightrun
def get_nr_info(dir):
    'Parses nr_info file into a dictionary that contains machine data.'
    nr_info_file = open(dir + "/nr_info.txt", 'r')
    global nr_info
    nr_info['checkboxes'] = {}
    checkboxes = [
        {'name': 'Cont. mode', 'grep': 'Continuous Mode'},
        {'name': 'MMD', 'grep': 'cvar_MMD_ENABLE'},
        {'name': 'MMC', 'grep': 'cvar_MMC_ENABLE'},
        {'name': 'DM', 'grep': 'cvar_DM_CREATE_IMAGE'},
        {'name': 'PI', 'grep': 'cvar_PERIPHERAL_ENABLE'},
        {'name': 'Registration', 'grep': 'cvar_CELL_REG_ENABLE'},
        {'name': 'OCR', 'grep': 'cvar_OCR_ENABLE'},
        {'name': 'DZ', 'grep': 'cvar_DETECTION_ZONES_ENABLE'},
        {'name': 'SVPI', 'grep': 'cvar_SVPI_ENABLE'},
        {'name': 'LSA', 'grep': 'cvar_LSA_ENABLE'},
        {'name': 'DMW', 'grep': 'cvar_DM_GENERAL_MURA_DETECTOR_ENABLE'},
    ]
    nr_info['scans'] = {}
    nr_info['qhr'] = {}
    nr_info['active_ipnodes'] = []
    # iterate over nr_info.txt
    for line in nr_info_file:
        for i in checkboxes:
            if i['grep'] in line:
                nr_info['checkboxes'][i['name']] = "true" in line.split()[-1]
        if "Current Version:" in line:
            nr_info['version'] = line.split()[-1]
            continue
        elif "Start Time:" in line:
            nr_info['start_time'] = line.split()[-1]
        elif "End Time:" in line:
            nr_info['end_time'] = line.split()[-1]
        elif line.startswith("|") and nr_info['scans'] == {}:
            try:
                nr_info['scans']['successfull'] = int(line.split()[5])
                nr_info['scans']['aborted'] = int(line.split()[7])
                nr_info['scans']['failed'] = int(line.split()[9])
            except:
                pass
        elif "Velocity" in line:
            nr_info['velocity'] = line.split()[-1]
        elif "Align Mode" in line:
            if line.split()[-1] == "PPAlign":
                nr_info['checkboxes']['PPAlign'] = True
                nr_info['checkboxes']['XIMAlign'] = False
            elif line.split()[-1] == "XIMAlign":
                nr_info['checkboxes']['PPAlign'] = False
                nr_info['checkboxes']['XIMAlign'] = True
        elif "AF Init Frequency" in line:
            nr_info['qhr']['AFInitFrequency'] = line.split()[-1]
        elif "Init Frequency" in line:
            nr_info['initIPFrequency'] = line.split()[-1]
        elif "DM DS rate" in line:
            nr_info['DM DS rate'] = line.split()[-1]
        elif "recipe:" in line:
            nr_info['recipe'] = line.split()[-1]
        elif "RIV mode" in line:
            nr_info['checkboxes']['VOF'] = 'VOF' in line.split()[-1]
            nr_info['checkboxes']['ET'] = 'ET' in line.split()[-1]
        elif "FVG:" in line:
            nr_info['checkboxes']['FVG'] = "true" in line.split()[-1]
        elif "MaxDefectImages" in line:
            nr_info['MaxDefectImages'] = line.split()[-1]
        elif "EndOfProcessingTimeout" in line:
            nr_info['endOfProcessingTimeout'] = line.split()[-1]
        elif "cvar_MAX_DEFECT_REPORTING_NUMBER " in line:
            nr_info['MAX_DEFECT_REPORTING_NUMBER'] = line.split()[-1].replace("'", "")
        elif "Grape Version" in line:
            nr_info['grape_v'] = line.split()[-1]
        elif "SPII s/w Version" in line:
            nr_info['spii_v'] = line.split()[-1]
        elif "CLB Version" in line:
            nr_info['clb_v'] = line.split()[-1]
        elif "QSIB Version" in line:
            nr_info['qsib_v'] = line.split()[-1]
        elif "Last rotation angle" in line:
            nr_info['rotat_angle'] = line.split()[-1]
        elif "MPI_Scheduler" in line and len(line.split()) == 5:
            nr_info['mpi_amount'] = line.split()[3]

        elif "HIP" in line:
            try:
                nr_info['active_ipnodes'].append(line.split()[0][-1])
                nr_info['cameras'].extend(line.split()[-1].split(','))
            except:
                nr_info['cameras'] = line.split()[-1].split(',')
    try:
        nr_info['cameras'] = map(int, nr_info['cameras'])
        nr_info['cameras'].sort()
        cam_data = {'consecutive': False, 'prev_cam': -1, 'first_cam': -1, 'cameras_readable': []}
        for cam in nr_info['cameras']:
            if cam_data['consecutive'] is False:
                if cam_data['prev_cam'] + 1 == cam:
                    cam_data['consecutive'] = True
                    cam_data['first_cam'] = cam_data['prev_cam']
                elif cam_data['prev_cam'] != -1:
                    cam_data['cameras_readable'].append(str(cam_data['prev_cam']) + ', ')
            else:
                if cam_data['prev_cam'] + 1 == cam:
                    cam_data['prev_cam'] = cam
                else:
                    cam_data['cameras_readable'].append(
                        str(cam_data['first_cam']) + "-" + str(cam_data['prev_cam']) + ', ')
                    cam_data['consecutive'] = False
            cam_data['prev_cam'] = cam
        if cam_data['consecutive'] is True:
            cam_data['cameras_readable'].append(str(cam_data['first_cam']) + "-" + str(cam_data['prev_cam']))
        else:
            cam_data['cameras_readable'].append(cam_data['prev_cam'])
        nr_info['cameras'] = ""
        for i in cam_data['cameras_readable']:
            nr_info['cameras'] += i
        # nr_info['cameras'] = cam_data['cameras_readable']
        nr_info['machine'] = dir.split('_')[-1]
        nr_info['timeframe'] = nr_info['start_time'] + "-" + nr_info['end_time']
    except:
        pass
    nr_info['date_of_analysis'] = date.today().strftime("%d/%m/%y")


# Create the mail with all the data
# @return - string that contains the e-mail HTML code
def create_mail():
    global exceptions
    stderr = "<b>There was a problem getting this information. please add manually.</b>"  # this will be added to mail where some inforamtion wasn't retrieved properly.
    checkbox_dir = (os.getcwd()).replace("\\", "/")
    get_nr_info(dir)
    csv_to_tables = [
        {
            'file': '/boot_rates.csv',
            'title': 'Boot Success Rates',
            'has_header': True,
            'convert_fracs_to_percentage': True,
            'digits_after_decimal_point': 2
        }
    ]
    print "getting nr_info - done"
    # executive summary template:
    mail_content = '<html><head><style>	h4 {text-decoration: underline} \n' \
                   'table {border-collapse:collapse; table-layout:fixed}\n' \
                   'td {border-style:solid; border-width:1px; word-wrap:break-word; white-space:nowrap}\n' \
                   '.bgGrey {background-color:grey}\n' \
                   'th {border-style:solid; border-width:1px; background-color:#4f81bd}</style>	</head>	<body>\n' \
                   '<h4>Executive Summary</h4><ul><li><u><b>Purpose:</b></u> <br>\n'
    try:
        mail_content += 'Test version ' + nr_info['version']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)

    # Create scan summary
    mail_content += '</li><li><u><b>Scan Summary</b></u><br>'
    try:
        mail_content += '<font color="green">' + str(
            nr_info['scans']['successfull']) + " scans were completed successfully. </font>"
        if nr_info['scans']['failed'] != 0:
            mail_content += "<br>\n"
            mail_content += '<font color="red">' + str(nr_info['scans']['failed']) + " scans failed. </font>"
        if nr_info['scans']['aborted'] != 0:
            mail_content += "<br>\n"
            mail_content += '<font color="red">' + str(nr_info['scans']['aborted']) + " scans aborted. </font>"
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)

    mail_content += '</li>\n<li><u><b>Main Issues</b></u></li>	' \
                    '<li><u><b>Detailed Problems and Analysis</b></u></li></ul>' \
                    '<h4>Logs: <a href="">Add log directory here</a></h4>\n'
    # creating table with machine name:
    mail_content += '<table class="tg"><tr><th colspan="5">' + dir.split('_')[-1] + '</th></tr>\n'
    # add date:
    mail_content += '<tr><td colspan="4">Date of Analyzing</td><td colspan="1">' + nr_info[
        'date_of_analysis'] + '</td></tr>\n'
    # add timeframe:
    mail_content += '<tr><td colspan="4">Time Frame</td><td colspan="1">'
    try:
        mail_content += nr_info['start_time'] + ' - ' + nr_info['end_time']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        mail_content += '</td></tr>\n'
    # add version:
    mail_content += '<tr><td colspan="4">Version</td><td colspan="1">'
    try:
        mail_content += nr_info['version']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        mail_content += '</td></tr>\n'
    # add recipe
    mail_content += '<tr><td colspan="4">Recipe</td><td colspan="1">'
    try:
        mail_content += nr_info['recipe']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        mail_content += '</td></tr>\n'
    # add grape version
    mail_content += '<tr><td colspan="4">Grape Version</td><td colspan="1">'
    try:
        mail_content += nr_info['grape_v']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        '</td></tr>\n'
    # add active nodes
    mail_content += '<tr><td colspan="4">Active Ipnodes</td><td colspan="1">'
    try:
        while len(nr_info['active_ipnodes']) > 0:
            mail_content += nr_info['active_ipnodes'].pop(0)
            if len(nr_info['active_ipnodes']) > 0:
                mail_content += ', '
    except:
        mail_content += stderr
    finally:
        mail_content += '</td></tr>\n'
    # add cameras:
    mail_content += '<tr><td colspan="4">Cameras</td><td colspan="1">'
    try:
        # for key in nr_info['cameras']:
        #     mail_content += key
        mail_content += nr_info['cameras']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    mail_content += '</td></tr>'
    # add stop reason
    mail_content += '<tr><td colspan="4">Stop Reason</td><td colspan="1">Stopped by User</td></tr>\n'
    # add velocity
    mail_content += '<tr><td colspan="4">Velocity</td><td colspan="1">'
    try:
        mail_content += nr_info['velocity']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        '</td></tr>\n'
    # add clb and qsib versions
    mail_content += '<tr><td colspan="4">CLB, QSIB versions</td><td colspan="1">'
    try:
        mail_content += nr_info['clb_v'] + nr_info['qsib_v']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        '</td></tr>\n'
    # add downsample rate
    mail_content += '<tr><td colspan="4">DM DS rate</td><td colspan="1">'
    try:
        mail_content += nr_info['DM DS rate']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        '</td></tr>'
    # add SPII Version
    mail_content += '<tr><td colspan="4">SPII Version</td><td colspan="1">'
    try:
        mail_content += nr_info['spii_v']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        '</td></tr>\n'
    # add Rotation Angle
    mail_content += '<tr><td colspan="4">Rotation Angle</td><td colspan="1">'
    try:
        mail_content += nr_info['rotat_angle']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        '</td></tr>\n'
    # add MPI_Schedulers
    mail_content += '<tr><td colspan="4">MPI Schedulers</td><td colspan="1">'
    try:
        mail_content += nr_info['mpi_amount']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        '</td></tr>\n'
    # add machine configuration headers
    mail_content += '<tr><th class="tg-031e" colspan="5">Machine Configuration</th></tr>' \
                    '<tr><td colspan="4"><b><u>Scan</b></u></td><td><b><u>Video</b></u></td></tr>\n<tr>'
    # add checkboxes and values
    i = 0
    video = ['VOF', 'ET', 'FVG']
    try:
        for key, value in nr_info['checkboxes'].iteritems():
            if key in ['VOF', 'ET', 'FVG']:
                continue
            elif value is True:
                mail_content += '<td colspan="1"><img src="' + checkbox_dir + '/checked.png"> ' + key + '</td>'
            else:
                mail_content += '<td colspan="1"><img src="' + checkbox_dir + '/unchecked.png"> ' + key + '</td>'
            i += 1
            if i == 4:
                mail_content += '<td colspan="1">'
                # add video
                if video:
                    cur = video.pop(0)
                    if nr_info['checkboxes'][cur] is True:
                        mail_content += '<img src="' + checkbox_dir + '/checked.png"> ' + cur
                    else:
                        mail_content += '<img src="' + checkbox_dir + '/unchecked.png"> ' + cur
                mail_content += '</td></tr><tr>'
                i = 0
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        if i < 4:
            mail_content += '<td colspan="' + str(4 - i) + '"></td>'
            i += 1
        mail_content += '<td></td></tr>'

    mail_content += '<tr><td colspan="4"><b><u>Parameter</b></u></td><td colspan="1"><b><u>Value</b></u></td></tr>\n'
    # add max defect reporting number
    mail_content += '<tr><td colspan="4">MAX_DEFECT_REPORTING_NUMBER</td><td colspan="1">'
    try:
        mail_content += nr_info['MAX_DEFECT_REPORTING_NUMBER']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        '</td></tr>\n'
    # add max defect reporting number with icons
    mail_content += '<tr><td colspan="4">MaxDefectImages</td><td colspan="1">'
    try:
        mail_content += nr_info['MaxDefectImages']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        '</td></tr>\n'
    # mail_content += '<tr><td>Max Defects</td><td>'+nr_info['Ma']
    # add InitIP frequency
    mail_content += '<tr><td colspan="4">Init IP Frequency</td><td colspan="1">'
    try:
        mail_content += nr_info['initIPFrequency']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        mail_content += '</td></tr>\n'
    # add endOfProcessingTimeout
    mail_content += '<tr><td colspan="4">End Of Processing Timeout</td><td colspan="1">'
    try:
        mail_content += nr_info['endOfProcessingTimeout']
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        mail_content += '</td></tr>\n'
    #add QHR information if needed
    if nr_info['qhr'] != {}:
        mail_content += '<tr><th class="tg-031e" colspan="5">QHR Configuration</th></tr>'
        # add AFInitFrequency
        mail_content += '<tr><td colspan="4">AF Init Frequency</td><td colspan="1">'
        try:
            mail_content += nr_info['qhr']['AFInitFrequency']
        except Exception as ex:
            mail_content += stderr
            exceptions.append(ex)
        finally:
            mail_content += '</td></tr>\n'

    # end of configuration table
    mail_content += '</table>\n'
    print "adding configuration table - done"

    # create max memory usage table
    mail_content += '<br><table><th>Process</th><th>Max Memory Usage (MB)</th>'
    try:
        mail_content += '<tr><td><b>AppExe</b></td><td>' + str(max_values['AppExe']) + '</td></tr>'
        mail_content += '<tr><td><b>GUIExec</b></td><td>' + str(max_values['GUIExec']) + '</td></tr>'
        mail_content += '<tr><td><b>MPI_Scheduler</b></td><td>' + str(max_values['MPI_Scheduler']) + '</td></tr>'
        mail_content += '<tr><td><b>CimsProxyExe</b></td><td>' + str(max_values['CimsProxyExe']) + '</td></tr>'
        mail_content += '<tr><td><b>VipExe</b></td><td>' + str(max_values['VipExe']) + '</td></tr>'
        for key in sorted(max_values):
            if key in ['AppExe', 'GUIExec', 'MPI_Scheduler', 'CimsProxyExe', 'VipExe']:
                pass
            else:
                mail_content += '<tr><td><b>' + str(key) + '</b></td><td>' + str(max_values[key]) + '</td></tr>'
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    finally:
        mail_content += '</table>\n'

    for i in csv_to_tables:
        i['file'] = dir + i['file']
        i_has_header = False
        i_convert_fracs_to_percentage = False
        i_digits_after_decimal_point = None
        if "has_header" in i:
            i_has_header = i['has_header']
        if "convert_fracs_to_percentage" in i:
            i_convert_fracs_to_percentage = i['convert_fracs_to_percentage']
        if "digits_after_decimal_point" in i:
            i_digits_after_decimal_point = i['digits_after_decimal_point']
        if "title" in i:
            mail_content += '<br><b><u>'
            mail_content += i['title']
            mail_content += '</b></u>\n'
        mail_content += '<br>\n'
        try:
            mail_content += csv_to_html_table(i['file'], i_has_header, convert_fracs_to_percentage=i_convert_fracs_to_percentage, digs_after_dec_point=i_digits_after_decimal_point)
        except Exception as ex:
            exceptions.append(ex)

    # create charts table
    try:
        for key in ['subsequences', 'memory', 'load']:
            # create each table seperately
            if key == "subsequences":
                mail_content += '<br><table><th colspan="2">Subsequences</th>\n'
            elif key == "memory":
                mail_content += '<br><table><th colspan="2">Memory Analysis</th>\n'
            else:
                mail_content += '<br><table><th colspan="2">Load Analysis</th>\n'
            # add charts for current table
            value = charts[key]
            while len(value) > 0:
                try:
                    mail_content += '<tr><td><img src="' + dir + '/chartImages/' + key + '/' + (value.pop(0)).replace(
                        " ", "_") + '.png"></td>\n'
                    if len(value) > 0: mail_content += '<td><img src="' + dir + '/chartImages/' + key + '/' + (
                        value.pop(0)).replace(" ", "_") + '.png"></td></tr>\n'
                except:
                    mail_content += stderr
            mail_content += '</table>\n'
            print "adding", key, "table - done"
    except Exception as ex:
        mail_content += stderr
        exceptions.append(ex)
    return mail_content

# Export chart images from Merged_charts.xlsx into chartImages directory
# @param dir - Nightrun directory
def export_graph_pictures(dir):
    global exceptions
    chartdir = dir + '/chartImages/'
    try:
        os.mkdir(chartdir)
    except Exception as ex:
        exceptions.append(ex)
    try:
        os.mkdir(chartdir + 'subsequences')
    except Exception as ex:
        exceptions.append(ex)
    try:
        os.mkdir(chartdir + 'load')
    except Exception as ex:
        exceptions.append(ex)
    try:
        os.mkdir(chartdir + 'memory')
    except Exception as ex:
        exceptions.append(ex)
    xlsApp = Dispatch("Excel.Application")
    xlsApp.Interactive = False
    xlsWB = xlsApp.Workbooks.Open(r'%s/Merged_charts.xlsx' % (dir))
    wbChartsSheet = xlsWB.Sheets(1)
    for chart in wbChartsSheet.ChartObjects():
        chart.Activate()
        chartTitle = chart.Chart.ChartTitle.Text
        # chart.Chart.ChartTitle.Text = chartTitle
        global charts
        current_chtdir = ""
        for key, value in charts.iteritems():
            if chartTitle in value:
                current_chtdir = chartdir + key + '/'
                chartTitle = chartTitle.replace(" ", "_")
        try:
            chart.Chart.Export(Filename=current_chtdir + chartTitle + '.png', FilterName="PNG")
        except Exception as ex:
            exceptions.append(ex)
    try:
        xlsWB.Close(True)
    except Exception as ex:
        exceptions.append(ex)
        try: xlsApp.quit()
        except Exception as ex: exceptions.append(ex)
    finally: xlsApp.Interactive = True


def parse_top(file):
    top_txt_file = open(file, 'r')
    to_parse = ("AppExe", "GUIExec", "MPI_Sched", "CimsProxy", "VipExe")  # modify if you need different data
    list = [['Time'], ['AppExe'], ['GUIExec'], ['CimsProxyExe'], ['MPI_Scheduler'], ['VipExe']]  # modify here too
    for line in top_txt_file:
        if line.startswith("top - "):
            if list[0][-1] == line.split()[2]: continue
            for i in list: i.append('0')
            list[0][-1] = line.split()[2]
        if any(i in line for i in to_parse):
            res = line.split()[5]
            # following code converts RES to Mb (if needed), and from string to float
            if res.endswith('m'):
                res = float(res[:-1])
            elif res.endswith('g'):
                res = float(res[:-1]) * 1000
            elif res.endswith('t'):
                res = float(res[:-1]) * 1000000
            else:
                try:
                    res = float(res) / 1000
                except:
                    pass

            # following code adds the RES to the list
            if line.split()[-1] in ["MPI_Sched+", "MPI_Schedu+"]:
                for i in list:
                    if i[0] == "MPI_Scheduler": i[-1] = res
            elif line.split()[-1].startswith('CimsPro'):
                for i in list:
                    if i[0] == "CimsProxyExe": i[-1] = res
            else:
                for i in list[1:]:
                    if line.split()[-1] == i[0]:
                        i[-1] = res
    global max_values
    for i in list[1:]:
        max_values[i[0]] = max((map(int, i[1:])))
    return list


def top_to_chart(file):  # new
    list = parse_top(dir + file)
    ActiveWS = mergedChartsWB.add_worksheet('Top Results')
    for i in range(len(list)):
        ActiveWS.write_column(0, i, list[i])
    # Following code creates the chart
    chartRow, chartCol = 0, 0
    global charts
    for i in range(1, len(list)):
        chart = mergedChartsWB.add_chart({'type': 'line'})
        chart.set_y_axis({'name': 'Used Memory (MB)'})
        chart.set_title({'name': list[i][0]})
        chart.add_series({'values': ['Top Results', 1, i, len(list[i]) - 1, i],
                          'categories': ['Top Results', 1, 0, len(list[i]) - 1, 0]})
        chart.set_legend({'position': 'none'})
        charts['memory'].append(list[i][0])
        chart_to_sheet(chart)
    print file, "done"


def csv_file_to_chart(file, sheet_name):
    global max_values
    with open(dir + file, 'r') as input_file:
        ActiveWS = mergedChartsWB.add_worksheet(sheet_name)
        row = 0
        max = 0
        global charts
        for line in input_file:
            line = line.split(',')
            match = re.search(r"\d\d:\d\d:\d\d",line[0])
            if match:
                line[0] = match.group(0)
            ActiveWS.write_row(row, 0, line)
            if file.startswith('/Load'):  # Load CSVs
                if len(line) != 6: continue
                if row == 0:
                    header, serieses = sheet_name, [line[1], line[2], line[3]]
                    charts['load'].append(header)
            else:  # free memory CSVs, defect number CSV
                if row == 0:
                    if file.startswith('/Free'):
                        if len(line) != 2: continue
                        # header (example) = Ipnode1 + Used memory without cache
                        header = file.split('/')[2].split('.')[0].capitalize() + line[1].replace('\n', '')
                        charts['memory'].append(header)
                    else:
                        header = sheet_name
                elif file.startswith('/Free'):
                    if int(line[1]) > max: max = int(line[1])
            row += 1

        chart = mergedChartsWB.add_chart({'type': 'line'})
        chart.set_title({'name': header})
        if file.startswith('/Load'):  # load files
            for i in range(1, 4):
                chart.add_series({'name': serieses[i - 1], 'values': [sheet_name, 1, i, row - 1, i],
                                  'categories': [sheet_name, 1, 0, row - 1, 0]})
                chart.set_y_axis({'name': 'Milliseconds'})
        else:
            chart.add_series({'values': [sheet_name, 1, 1, row - 1, 1],
                              'categories': [sheet_name, 1, 0, row - 1, 0]})
            chart.set_legend({'position': 'none'})
            if file.startswith("/Defect") or file.startswith("/Macro_Defects"):
                chart.set_y_axis({'name': 'Amount'})
            else:
                chart.set_y_axis({'name': 'Used Memory (MB)'})
                max_values[file.split('/')[2].split('.')[0].capitalize()] = max  # example: max_values['Host'] = 6000
        chart_to_sheet(chart)
        print file, "done"


def tact_to_chart(file):
    with open(dir + file, 'r') as input_file:
        ActiveWS = mergedChartsWB.add_worksheet('Tact Results')
        row = 0
        for line in input_file:
            line = line.split(',')
            if row == 0:
                headers = line
                headers[-1] = headers[-1].strip()  # omit \n from last header
                global charts
                charts['subsequences'].extend(headers[1:])
                ActiveWS.write_row(row, 0, line)
            else:
                currentTime = line.pop(0).split(':')
                currentTime = time(int(currentTime[0]), int(currentTime[1]), int(currentTime[2][:2]))
                if row == 1: start_time = currentTime
                ActiveWS.write(row, 0, currentTime, time_format)
                ActiveWS.write_row(row, 1, line)
            row += 1
        for i in range(1, len(headers)):
            chart = mergedChartsWB.add_chart({'type': 'line'})
            chart.set_title({'name': headers[i]})
            chart.add_series({'categories': ['Tact Results', 1, 0, row - 1, 0],
                              'values': ['Tact Results', 1, i, row - 1, i]})
            chart.set_legend({'position': 'none'})
            chart.set_y_axis({'name': 'Seconds'})
            chart_to_sheet(chart)
        print file, "done"


def find_file(file_start):
    path = os.listdir(dir)
    for i in path:
        if i.startswith(file_start):
            return '/' + i
    return file_start


def chart_to_sheet(chart):  # Adds an existing chart to the charts sheet
    global chartRow
    global chartCol
    mergedChartsWS.insert_chart(chartRow, chartCol, chart)
    if chartCol != 8:
        chartCol = 8
    else:
        chartRow, chartCol = chartRow + 15, 0


def update_nr_archive():
    global exceptions
    try:
        nr_archive_path = r"\\main_w\vol2\USR_DATA\UNIT\DISPLAY\Groups\AOI\qa\qtm\tests\\" + nr_info['version'][
                                                                                             :nr_info['version'].rfind(
                                                                                                 ".")] + "\NightRun_Scans_Archive.xlsx"
        list_to_append = [nr_info['date_of_analysis'],
                          nr_info['machine'],
                          nr_info['timeframe'],
                          nr_info['version'],
                          nr_info['recipe'],
                          nr_info['cameras'],
                          nr_info['velocity'],
                          "Stopped By User",
                          nr_info['scans']['successfull'],
                          nr_info['scans']['aborted'],
                          nr_info['scans']['failed'],
                          str(nr_info['checkboxes']['SVPI']),
                          str(nr_info['checkboxes']['PPAlign']),
                          str(nr_info['checkboxes']['DZ']),
                          str(nr_info['checkboxes']['Registration']),
                          str(nr_info['checkboxes']['MMC']),
                          str(nr_info['checkboxes']['MMD']),
                          str(nr_info['checkboxes']['DM']),
                          str(nr_info['checkboxes']['DMW']),
                          str(nr_info['checkboxes']['PI']),
                          str(nr_info['checkboxes']['LSA']),
                          str(nr_info['checkboxes']['ET']),
                          str(nr_info['checkboxes']['VOF']),
                          str(nr_info['checkboxes']['FVG']),
                          str(nr_info['checkboxes']['OCR'])]
        nr_archive = openpyxl.load_workbook(nr_archive_path)
        rel_sheet = nr_archive.active
        rel_sheet.append(list_to_append)
        nr_archive.save(nr_archive_path)
        print "NR archive was appended successfully."
    except Exception as ex:
        exceptions.append(ex)


def parse_syslog_events():
    pass


###### NEEDED OBJECTS ######
a = Tk()
a.withdraw()
dir = askdirectory()
exceptions = []
charts = {'subsequences': ['Defects Number', 'Macro Defects Number'], 'memory': [], 'load': []}
mergedChartsWB = xlsxwriter.Workbook((dir + '/Merged_charts.xlsx'), {'strings_to_numbers': True})
mergedChartsWS = mergedChartsWB.add_worksheet('Merged Charts')
time_format = mergedChartsWB.add_format({'num_format': 'hh:mm:ss'})
chartRow, chartCol = 0, 0
max_values = {}
nr_info = dict()
html_mail = ""

###### MAIN PROGRAM ######
try:
    csv_files = [find_file('Defect'), find_file('Macro_Defects')]
    sheet_names = ['Defects Number', 'Macro Defects Number']
except Exception as ex:
    exceptions.append(ex)
try:
    for i in os.listdir(dir + '/Free_Mem'):
        try:
            csv_files.append('/Free_Mem/' + i)
            sheet_names.append(i.split('.')[0].capitalize() + " free mem")
        except Exception as ex:
            exceptions.append(ex)
except Exception as ex:
    exceptions.append(ex)
try:
    for i in os.listdir(dir + find_file('Load')):
        if i.startswith('load'):
            csv_files.append(find_file('Load') + '/' + i)
            if i.startswith('load_0'):
                sheet_names.append('Host load')
            else:
                sheet_names.append('Ipnode ' + i[5] + ' load')
except Exception as ex:
    exceptions.append(ex)
try:
    tact_to_chart(find_file('Tact'))
except Exception as ex:
    exceptions.append(ex)
try:
    top_to_chart('/CPU_Test/host.txt')
except Exception as ex:
    exceptions.append(ex)
for file in csv_files:
    try:
        csv_file_to_chart(file, sheet_names.pop(0))
    except Exception as ex:
        exceptions.append(ex)

mergedChartsWB.close()
print "Merged_charts file was created.\nExporting chart images..."
export_graph_pictures(dir)
print "Creating mail..."
mail = create_mail()
mail_file = open(dir + '/mail.html', 'w+')
mail_file.write(mail)
mail_file.flush()

print "done"

if raw_input("Update NR archive? Input 'Y' to approve\n").upper() in ['Y', 'YES']:
    update_nr_archive()

if exceptions != []:
    print "Please note following exceptions:"
    for i in exceptions:
        print i
    raw_input("press [Enter] to exit")
