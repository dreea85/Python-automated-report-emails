import pandas as pd
import os
import win32com.client as win32
from datetime import date, datetime, timedelta
import time as time_module
import openpyxl
import csv
import logging

logging.basicConfig(filename='C:\\path\email_log.log', level=logging.DEBUG)
logging.info('Script started')

outlook_mail = win32.Dispatch("Outlook.Application")


# track the number of emails sent:
def log_email(recipient, subject):
    with open('sent_emails.csv', 'a', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([recipient, subject, date.today().strftime('%Y-%m-%d')])


# Construct email body
msg_body_beginning = '''
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>Hi {first_name},</p>
    <br>
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>text here</p><p style='color:black;font-size:16px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>text here</p>
    <br>
    '''

msg_body_table1 = '''
    <table border = "1" style="border-collapse: collapse;width:70%;font-size:15px;font-family:'Times New Roman'">
        <tr style="font-weight: bold;">
            <td style="width:50%;padding: 10px; background-color: paleturquoise">text</td>
            <td style="padding: 10px; background-color: paleturquoise">text</td>
        </tr>
        <tr>
            <td style="width:60%;padding: 10px;">number</td>
            <td style="padding: 10px;">{number}</td>
        </tr>
        <tr>
            <td style="width:60%;padding: 10px;">Reviewer ID</td>
            <td style="padding: 10px;">{user_id}</td>
        </tr>
        <tr>
            <td style="width:60%;padding: 10px;">text</td>
            <td style="padding: 10px;">{review_timestamp}</td>
        </tr>
        <tr>
            <td style="width:60%;padding: 10px;">text</td>
            <td style="padding: 10px;">{defect_type}</td>
        </tr>
    </table>
    '''

msg_body_table2 = '''
        </tbody>
    </table>
    '''
msg_body_end = '''
    <br>
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>
        <strong><span style="color:black;">text</span></strong>
        <br>
        <span style="color:black;">&bull;text</span>
        <br>   
        <span style="color:black;">&bull;text</span>
    </p>
    <br>
    <p><strong><span style='font-size:15px;font-family:"Times New Roman",serif;color:black;'>Thank you in advance for your understanding and cooperation!</span></strong></p>
    <br>
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'><br></p>
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>Best regards,<br></p>
    <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>Reporting Team</p>
    '''


def sendMailtoReviewer(number, user_id, first_name, review_timestamp, week, defect_type, excel_file_path):
    print(f"Preparing the email")

    mail = outlook_mail.CreateItem(0)

    # mail.To = f'{user_id}@domain.com'
    mail.To = "email@domain.com"
    mail.SentOnBehalfOfName = "automation_team@domain.com"
    mail.Subject = f'Record from week {week}'
    mail.HTMLBody = (msg_body_beginning.format(first_name=first_name) +
                     msg_body_table1.format(user_id=user_id,
                                            number=number,
                                            review_timestamp=review_timestamp,
                                            defect_type=defect_type) +
                     msg_body_table2 +
                     msg_body_end)

    attachment = excel_file_path
    mail.Attachments.Add(attachment)
    # mail.Display()
    mail.Send()

def aggregate_defect_types(files):
    print(f"Processing file: {files}")
    all_data = pd.DataFrame()
    for file_path in files:
        data = pd.read_csv(file_path, sep='\t')

def convert_to_excel(txt_file_path):
    data = pd.read_csv(txt_file_path, sep='\t')
    excel_file_path = txt_file_path.replace('.txt', '.xlsx')
    data.to_excel(excel_file_path, index=False)
    return excel_file_path
# process new file function
def process_new_file(file_path):
    print(f"Processing file: {file_path}")
    data = pd.read_csv(file_path, sep='\t')

    # print("Original columns are:", data.columns)
    data.columns = [col[1:] if col.strip().startswith('t') else col for col in data.columns]
    # print("Adjusted columns are:", data.columns)

    # split the time, take only date without the hour
    data['review_timestamp'] = data['review_timestamp'].astype(str).str.split().str[0]
    data['review_timestamp'] = pd.to_datetime(data['review_timestamp'])
    data['review_timestamp'] = data['review_timestamp'].dt.strftime("%b-%d-%Y")

    if data.empty:
        print('The dataframe is empty. No data to process.')

    excel_file_path = convert_to_excel(file_path)

    for index, row in data.iterrows():
        sendMailtoReviewer(row['number'], row['user_id'], row['first_name'], row['review_timestamp'], row['week'],
                           row['defect_type'], excel_file_path)

    return excel_file_path


def aggregate_data_previous_week(path):
    print(f"Processing file: {path}")
    all_data_frames = []
    for i in range(1,7):    #loop over the last 6 days
        date_to_check = date.today() - timedelta(days=i)
        file_name = f"text {date_to_check.strftime('%Y-%m-%d')}.txt"
        file_path = os.path.join(path, file_name)
        if os.path.exists(file_path):
            daily_data = pd.read_csv(file_path, sep='\t')
            all_data_frames.append(daily_data)

    if all_data_frames:
        aggregated_data = pd.concat(all_data_frames, ignore_index=True)
    else:
        aggregated_data = pd.DataFrame()  #Return an empty Dataframe if no data was found

    return aggregated_data


def get_defect_type_counts(aggregated_data):

    pivot_data = aggregated_data.pivot_table(
        index='defect_type',
        columns='node_name',
        aggfunc='size',
        fill_value=0)
    # pivot_data.columns.name = None
    # pivot_data.loc['Total'] =pivot_data.sum()
    # defect_counts = aggregated_data.groupby(['defect_type', 'node_name']).size().reset_index(name='Count')
    # defect_counts.columns = ['Defect Type', 'Count']
    print(pivot_data)
    return pivot_data


def get_previous_week_number():
    today = date.today()
    one_week_ago = today - timedelta(weeks=1)
    iso_calendar = one_week_ago.isocalendar()
    week_number = iso_calendar[1]
    return week_number


def email_to_manager(defect_counts, week_number):
    #convert defect_counts dataframe to html table for email body
    html_table = defect_counts.to_html(index=False)
    manager_email_body = msg_body_beginning2 + html_table + msg_body_end2
    return manager_email_body

msg_body_beginning2 = f'''
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>Hello, all</p>
                <br>
                <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>text Week {get_previous_week_number()} :</p><p style='color:black;font-size:16px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'></p>
                <br>
                '''

msg_body_end2 = '''
            <br>
            <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>
                <strong><span style="color:black;">text here</span></strong>
                <br>
                <span style="color:black;"><a href="https://path" target="_blank">text</a></span>
                <br>   
                <span style="color:black;"><a href="https://path">text</a></span>
            </p>
            <br>

            <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'><br></p>
            <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>Best regards,<br></p>
            <p style='margin-right:0in;margin-left:0in;font-size:15px;font-family:"Times New Roman",serif;margin:0in;margin-bottom:.0001pt;'>Reporting Team</p>
            '''

def create_summary_table(pivot_data):
    header_color = "#48d1cc"
    row_color = "#f2f2f2"
    total_color = "#eac4ad"

    html_table = '<table border = "1" style="border-collapse: collapse;width:70%;font-size:15px;font-family:\'Times New Roman\'">'
    #add header row
    html_table += f'<tr style="background-color: {header_color}; color: black;">'
    html_table += '<td>Defect Type</td>'
    for node in pivot_data.columns:
        html_table += f'<td>{node}</td>'
    html_table += '<td>Total</td><tr>'


    #add data rows
    for idx, (defect, row) in enumerate(pivot_data.iterrows()):
        bg_color = row_color if idx % 2 == 0 else "white"
        html_table += f'<tr style="background-color: {bg_color};">'
        html_table += f'<td>{defect}</td>'
        row_sum = 0
        for count in row:
            html_table += f'<td>{count}</td>'
            row_sum += count
        html_table += f'<td>{row_sum}</td></tr>'

    column_total = pivot_data.sum()
    html_table += f'<tr style="background-color: {total_color}; font-weight: bold;">'
    html_table += '<td>Total</td>'
    grand_total = 0
    for total in column_total:
        html_table += f'<td>{total}</td>'
        grand_total += total
    html_table += f'<td>{grand_total}</td></tr>'

    html_table += '</table>'
    return html_table


def sendMailtoManager(emails, summary_table_html):
    print(f"Preparing the email to the managers")

    mail = outlook_mail.CreateItem(0)

    # mail.To = f'{user_id}@domain.com'
    mail.To = "email@domain.com"
    mail.SentOnBehalfOfName = "automation_team@domain.com"
    mail.Subject = f'text- Week {get_previous_week_number()} '

    #we use the defect_counts dataframe directly here
    summary_table_html = create_summary_table(pivot_table_data)

    mail.HTMLBody = (msg_body_beginning2 +
                    summary_table_html +
                    msg_body_end2)

    # mail.Display()
    mail.Send()


def check_for_previous_file(path):
    # substract 1 day from today's date
    previous_day = date.today() - timedelta(days=1)
    # format the date
    file_name = f"text  {previous_day.strftime('%Y-%m-%d')}.txt"
    file_path = os.path.join(path, file_name)
    if os.path.exists(file_path):
        print(f"Found today's file: {file_name}")
        process_new_file(file_path)
    else:
        print(f"Today's file {file_name} not found")


logging.info('Email sent')

if __name__ == "__main__":
    print("Script started.")
    path = r'\\path folder'

    # check for today file
    check_for_previous_file(path)
    previous_week_number =get_previous_week_number()
    print(f'the previous iso week number is {previous_week_number}')

    if datetime.today().weekday() == 3:    #0 is Monday, 6 is Sunday ecc

        aggregated_data = aggregate_data_previous_week(path)
        pivot_table_data = get_defect_type_counts(aggregated_data)
        summary_table_html = create_summary_table(pivot_table_data)
        # manager_email_body = email_to_manager(defect_counts, previous_week_number)
        sendMailtoManager('email@domain.com', summary_table_html)

    print('Script finished')
