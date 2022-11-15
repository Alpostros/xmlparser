from gettext import find
import xml.etree.cElementTree as et
import pandas as pd
import sys
import os

def get_test_report_type(filename):
    s = ""
    if filename.find('USSD') !=  -1:
        s = filename[filename.find('USSD'):filename.find('USSD')+4]

    elif filename.find('API') != -1:
        s = filename[filename.find('API'):filename.find('API')+3]

    else:
        s = -1
    
    return s

final_df = pd.DataFrame()
filenames = os.listdir('C:\\Users\\eozealp\\Desktop\\ParseXML\\')
for filename in filenames:
    test_report_type =  get_test_report_type(filename)
    if test_report_type == -1:
        print("Test report type (API/USSD) unknown in filename, skipping for file: "+ filename)
        continue

    tree = et.parse(filename)
    root = tree.getroot()
    opco_name = filename[:filename.find('-')]
    dt = root.attrib['generated']
    test_report_date = dt[:8]
    opco = []
    test_type = []
    test_name = []
    status = []
    reason = []
    resp = []
    request = []
    error_codes = ["HTTP/1.1 401 Unauthorized", "HTTP/1.1 500 Server Error", "HTTP/1.1 302 Found", "HTTP/1.1 404 Not Found", "HTTP/1.1 403 Forbidden", "No route to host", "Connection refused"] # 302 is not an error code, it is a response that shows OTP is Activated
    for reg in root.iter('test'):
        root1 = et.Element('root')
        root1 = reg
        if reg[-1].attrib['status'] == 'FAIL':
            opco.append(opco_name)
            test_name.append(reg.attrib['name'])
            status.append(reg[-1].attrib['status'])
            for child in reg.iter():   
                if child.tag=="kw" and test_report_type=='API':
                    request_string = "Request string not found."
                    for child2 in child.findall("./kw[@name='Log']"):
                        for child3 in child2.findall("./msg"):
                            id1 = child3.text.find("<?xml")
                            if id1 != -1:
                                id2 = len(child3.text)
                                request_string = child3.text[id1:id2]
                                request.append(request_string)
                        break;
                    

                if child.tag == "status" and child.text != None:
                    reason_string = child.text
                    if test_report_type == "USSD":
                        test_type.append("USSD")
                        if reason_string.find('applicationResponse') != -1:
                            id1 = reason_string.find('applicationResponse')
                            id2 = reason_string.find('</applicationResponse>')
                            reason.append(reason_string[id1+20:id2])
                        else:
                            reason.append("Application Response is Empty!")

                        id1 = reason_string.find('</response')
                        id2 = len(reason_string)
                        resp.append(reason_string)
                        request.append("Request strings are not checked for USSD.")
                        
                    elif test_report_type == "API":
                        test_type.append("API")
                        error_name = "No error code found"
                        for err in error_codes:
                            if reason_string.find(err) != -1:
                                if err == "HTTP/1.1 302 Found":
                                    error_name = "OTP Activated"
                                else:
                                    error_name = err
                    
                        reason.append(error_name)
                        resp.append(reason_string)

        elif reg[-1].attrib['status'] == 'PASS' or reg[-1].attrib['status'] == 'SKIP':
            opco.append(opco_name)
            test_name.append(reg.attrib['name'])
            status.append(reg[-1].attrib['status'])
            request.append("")
            reason.append("")
            resp.append("")
            if test_report_type == "USSD":
                test_type.append("USSD")
            elif test_report_type == "API":
                test_type.append("API")


                        

                        
    d = {'Opco Name': opco,'Test Type':test_type ,'Test Name': test_name, 'Status': status, 'Root Cause (ApplicationResponse)': reason, 'Full Response': resp, 'Request': request}                    
    df = pd.DataFrame.from_dict(d, orient='index')
    df = df.transpose()
    final_df = pd.concat([final_df, df], axis=0)

no_error_list = ["Application Response is Empty!", "No error code found"]
df_errors = final_df[final_df["Root Cause (ApplicationResponse)"].isin(no_error_list)]
writer = pd.ExcelWriter('mtn-failed-cases-'+test_report_date+'.xlsx')

final_df.to_excel(writer, sheet_name="Failed Cases")
df_errors.to_excel(writer, sheet_name = "Imp. Errors")
writer.save()
