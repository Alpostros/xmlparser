from gettext import find
import xml.etree.cElementTree as et
import pandas as pd
import sys
import os
from datetime import datetime

def get_test_report_type(filename):
    s = ""
    if filename.find('USSD') !=  -1:
        s = filename[filename.find('USSD'):filename.find('USSD')+4]

    elif filename.find('API') != -1:
        s = filename[filename.find('API'):filename.find('API')+3]

    else:
        s = -1
    return s

def get_elapsed_time(root, elapsed_time):
    starttime = root[0][1].attrib['starttime']
    endtime = root[0][1].attrib['endtime']
    starttime_formatted = datetime.strptime(starttime[9:], "%H:%M:%S.%f")
    endtime_formatted = datetime.strptime(endtime[9:], "%H:%M:%S.%f")
    execution_time_api = endtime_formatted - starttime_formatted
    execution_time_api = str(execution_time_api)[:7]
    elapsed_time.append(execution_time_api)

def get_request_response(request, tag_kw, resp, reason, test_report_type, status):
    request_string = "Request string not found."
    reason_string = ""
    login_req_string = ""
    for child2 in tag_kw:
        for child3 in child2.findall("./msg"):
            id1 = child3.text.find("<?xml")
            if id1 != -1 and "response" not in child3.text and "Response" not in child3.text: # Only applicable to API.
                if "p:loginrequest" not in child3.text:
                    id2 = len(child3.text)
                    request_string = child3.text[id1:id2]
                elif "p:loginrequest" in child3.text:
                    id2 = len(child3.text)
                    login_req_string = child3.text[id1:id2]
                    
    if status == "FAIL": # To get the response and define the reason of failed cases, applies to both cases.
        for child3 in child2.findall("./msg"):
            reason_string = child3.text
    
    if status == "FAIL":
        if test_report_type == "API":
            api_part(reason_string, resp, reason)
        else:
            ussd_part(reason_string, resp, reason)
    else:
        resp.append("")
        reason.append("")

    if request_string == "Request string not found." and len(login_req_string) > 0: # For login request, if there is only one of it
        request.append(login_req_string)
        return

    elif test_report_type != "USSD": # Request strings are only applicable to API.
        request.append(request_string)
        return
                                              
def ussd_part(reason_string, resp, reason):
    if reason_string != None:
        if reason_string.find('applicationResponse') != -1:
            id1 = reason_string.find('applicationResponse')
            id2 = reason_string.find('</applicationResponse>')
            reason.append(reason_string[id1+20:id2])
        else:
            reason.append("Application Response is Empty!")

        id1 = reason_string.find('</response')
        id2 = len(reason_string)
        resp.append(reason_string)

def api_part(reason_string, resp, reason):
    error_codes = ["HTTP/1.1 401 Unauthorized", "HTTP/1.1 500 Server Error", "HTTP/1.1 302 Found", "HTTP/1.1 404 Not Found", "HTTP/1.1 403 Forbidden", "No route to host", "Connection refused"] # 302 is not an error code, it is a response that shows OTP is Activated
    error_name = "No error code found"
    for err in error_codes:
        if reason_string == None:
            continue
        elif reason_string.find(err) != -1:
            if err == "HTTP/1.1 302 Found":
                error_name = "OTP Activated"
            else:
                error_name = err

    reason.append(error_name)
    resp.append(reason_string)
    
final_df = pd.DataFrame()
dir_path = os.path.dirname(os.path.realpath(__file__))
filenames = os.listdir(dir_path)
for filename in filenames:
    test_report_type =  get_test_report_type(filename)
    if test_report_type == -1:
        print("Test report type (API/USSD) unknown in filename, skipping for file: "+ filename)
        continue

    tree = et.parse(dir_path+"/"+filename)
    root = tree.getroot()
    opco_name = filename.split('-')[0]
    dt = root.attrib['generated']
    test_report_date = dt[:8]
    opco = []
    test_type = []
    test_name = []
    status = []
    reason = []
    resp = []
    request = []
    elapsed_time= []
    for reg in root.iter('test'):
        root1 = et.Element('root')
        root1 = reg
        opco.append(opco_name)
        test_name.append(reg.attrib['name'])
        status.append(reg[-1].attrib['status'])
        get_elapsed_time(root, elapsed_time)
        test_type.append(test_report_type)
        if test_report_type == "USSD":
            request.append("Requests are not checked for USSD.")

        kw1 = reg.findall(".//kw/kw[@name='Log']")
        get_request_response(request, kw1, resp, reason, test_report_type, reg[-1].attrib['status'])

        
    d = {'Opco Name': opco,'Test Type':test_type ,'Test Name': test_name, 'Status': status, 'Root Cause (ApplicationResponse)': reason, 'Full Response': resp, 'Request': request, 'Total Elapsed Time': elapsed_time}                    
    df = pd.DataFrame.from_dict(d, orient='index')
    df = df.transpose()
    final_df = pd.concat([final_df, df], axis=0)

no_error_list = ["Application Response is Empty!", "No error code found"]
writer = pd.ExcelWriter(dir_path+'/mtn-failed-cases-'+test_report_date+'.xlsx')
final_df.to_excel(writer, sheet_name="Failed Cases")
writer.save()
