from docx.shared import Inches
from docx import Document
from docx.shared import Pt
from modified_report_script import ETL
import datetime as dt

def add_dataframe_to_doc(doc, dataframe):
    # Add a table to the document for the given dataframe
    table = doc.add_table(rows=1, cols=len(dataframe.columns))
    # style
    table.style = 'Light Shading Accent 1'
    table.autofit = True
    hdr_cells = table.rows[0].cells
    for i, column_name in enumerate(dataframe.columns):
        try:
            mod_col_name = column_name.replace('_', ' ').capitalize()
        except AttributeError:
            mod_col_name = column_name
        hdr_cells[i].text = str(mod_col_name)
        # set font size and make them bold
        hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(10)
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
    for _, row in dataframe.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            try:
                value = round(value, 2)
            except:
                pass
            row_cells[i].text = str(value)
            #set font size
            row_cells[i].paragraphs[0].runs[0].font.size = Pt(8)

date_from = '2024-01-08 00:00:00'
date_to = '2024-01-11 23:59:59'
try:
    answered_df, missed_df,  error_calls, agents, agents_ext, unanswered_call_counts, unanswered_calls_ring, unanswered_calls_hour = ETL(date_from, date_to)
except Exception as e:
    print(e)
    exit(1)
doc = Document()
# modify document margins
sections = doc.sections
for section in sections:
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    
doc.add_heading('CloudTalk - Mitsis Group Analytics', 0)
doc.add_heading("Date range: " + dt.datetime.strptime(date_from, '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y %H:%M:%S') + " - " + dt.datetime.strptime(date_to, '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y %H:%M:%S'), level=2)
doc.add_paragraph('This report was generated on ' + dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S') + ' automatically by the Mitsis Group Business Analytics team.\nNote that all times are depicted in "seconds".')

doc.add_heading('General Statistics', level=4)
p1 = doc.add_paragraph('Total Answered calls retrieved:')
p1.add_run(f' {len(answered_df)}').bold = True
p2 = doc.add_paragraph('Total Missed calls retrieved:')
p2.add_run(f' {len(missed_df)}').bold = True
p3 = doc.add_paragraph('Total Calls that failed to be retrieved through API:')
p3.add_run(f' {error_calls}').bold = True
# top agent in terms of total answered calls (retrivied from "agents" dataframe)
top_agent = agents[agents['Answered_Calls'] == agents['Answered_Calls'].max()]['agent_name'].values[0]
# agent with most unanswered calls (retrivied from "unanswered_call_counts" dataframe)
top_unanswered_agent = unanswered_call_counts[unanswered_call_counts['Unanswered_UniqCount'] == unanswered_call_counts['Unanswered_UniqCount'].max()]['agent_name'].values[0]
# hour with most unanswered calls (retrivied from "unanswered_calls_hour" dataframe)
# add a "Total" row to the dataframe
unanswered_calls_hour_cp = unanswered_calls_hour.copy()
unanswered_calls_hour_cp.loc['Total'] = unanswered_calls_hour_cp.sum()
# find the column (hour) with the most unanswered calls
top_hour = unanswered_calls_hour_cp.loc['Total'].idxmax()
del unanswered_calls_hour_cp
# format the hour from an int to a time string
top_hour = dt.datetime.strptime(str(top_hour), '%H').strftime('%H:%M')
 
p4 = doc.add_paragraph('Top agent in terms of total ')
p4.add_run('answered').bold = True
p4.add_run(' calls:')
p4.add_run(f' {top_agent}').bold = True
p5 = doc.add_paragraph('Top agent/site in terms of total ')
p5.add_run('unanswered').bold = True
p5.add_run(' calls:')
p5.add_run(f' {top_unanswered_agent}').bold = True
p6 = doc.add_paragraph('Top hour in terms of most ')
p6.add_run('unanswered').bold = True
p6.add_run(' calls:')
p6.add_run(f' {top_hour}').bold = True
p1.style = p2.style = p3.style = p4.style = p5.style = p6.style = 'List Bullet'
doc.add_heading('1. Agent Statistics (incl. Internal Calls)', level=1)
add_dataframe_to_doc(doc, agents)

doc.add_heading('2. Agent Statistics (excl. Internal Calls)', level=1)
add_dataframe_to_doc(doc, agents_ext)

doc.add_heading('3. Unanswered unique number calls per agent (except internal)', level=1)
add_dataframe_to_doc(doc, unanswered_call_counts)

doc.add_heading('4. Distribution of unanswered calls per ring time (0-10s, 10-20s, etc.)', level=1)
#add_dataframe_to_doc(doc, unanswered_calls_ring)
doc.add_picture('unanswered_calls_ring_plot.png', width=Inches(7.6), height=Inches(2.5))

doc.add_heading('5. Distribution of unanswered calls per hour and agent', level=1)
unanswered_calls_hour['Total'] = unanswered_calls_hour.sum(axis=1)
add_dataframe_to_doc(doc, unanswered_calls_hour.reset_index())
doc.add_heading('6. Distribution of unanswered calls per hour and agent (Plots)', level=1)
doc.add_picture('unanswered_calls_hour_plot.png', width=Inches(7.6), height=Inches(2))
doc.add_picture('hours.png', width=Inches(7.6), height=Inches(5))


doc.save('Cloudtalk_Report.docx')


