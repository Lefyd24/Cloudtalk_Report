import streamlit as st
import requests
import datetime as dt
import pandas as pd
import matplotlib.pyplot as plt
import altair as alt
from dotenv import load_dotenv, find_dotenv
import os
from warnings import filterwarnings
from handle_data import edit_data
filterwarnings('ignore')
load_dotenv()

def create_altair_chart(agent_name, data):
    chart = alt.Chart(data).mark_bar().encode(
        x='Hour:N',
        y='Values:Q',
        tooltip=['Hour', 'Values']
    ).properties(
        title=f"{agent_name}",
        width=250,
        height=300,
        
    )
    return chart


def call_api(date_from, date_to):
    idd = os.environ.get('API_USER')
    pw = os.environ.get('API_PW')
    # Make API call here
    api_url = "https://my.cloudtalk.io/api/calls/index.json"
    page = 1
    limit = 1000
    answered_df = pd.DataFrame()
    missed_df = pd.DataFrame()
    
    if date_from > date_to:
        st.error("Error: Start date is after end date")
        return

    while True:
        print("Page: " + str(page))
        # make requests
        answered = requests.get(api_url, auth=(idd, pw), params={'page': page, 'limit': limit, 'date_from': date_from, 'date_to': date_to, 'status': 'answered'})
        page_count = answered.json()['responseData']['pageCount']
        
        #print(r.json())
        if answered.status_code != 200:
            print("Error: " + str(answered.status_code))
            break
        # convert to dataframe
        answered_df = pd.concat([answered_df, pd.DataFrame(answered.json()['responseData']['data'])])
        page += 1
        if page >= page_count:
            break

    page = 1
    while True:
        print("Page: " + str(page))
        # make requests
        missed = requests.get(api_url, auth=(idd, pw), params={'page': page, 'limit': limit, 'date_from': date_from, 'date_to': date_to, 'status': 'missed'})
        
        page_count = missed.json()['responseData']['pageCount']
        
        #print(r.json())
        if missed.status_code != 200:
            print("Error: " + str(missed.status_code))
            break
        # convert to dataframe
        missed_df = pd.concat([missed_df, pd.DataFrame(missed.json()['responseData']['data'])])
        page += 1
        if page >= page_count:
            break
        
    all_call_info = edit_data(answered_df, missed_df)

    # Stats per agent
    all_call_info['talking_time'] = all_call_info['talking_time'].astype(float)
    all_call_info['waiting_time'] = all_call_info['waiting_time'].astype(float)
    all_call_info['total_time'] = all_call_info['total_time'].astype(float)
    agents = all_call_info.groupby(['agent_id', 'agent_name']).agg(Calls=('id', 'count'), Avg_Talking_Time=('talking_time', 'mean'), Avg_Waiting_Time=('waiting_time', 'mean'), Avg_Total_Time=('total_time', 'mean'), 
                                                                   Total_Talking_Time=('talking_time', 'sum'), Total_Waiting_Time=('waiting_time', 'sum'), Total_Time=('total_time', 'sum')
                                                                   ).reset_index()
    # Crosstab to count 'missed' and 'answered' calls
    status_counts = pd.crosstab(index=[all_call_info['agent_id'], all_call_info['agent_name']], columns=all_call_info['status'])

    # Merge the two DataFrames
    agents = pd.merge(agents, status_counts, how='left', left_on=['agent_id', 'agent_name'], right_index=True)

    # If you want to rename the status columns for clarity
    agents.rename(columns={'answered': 'Answered_Calls', 'missed': 'Missed_Calls'}, inplace=True)
    #agents.rename(columns={'id': 'calls', 'talking_time': 'avg_talking_time', 'waiting_time': 'avg_waiting_time', 'total_time': 'avg_total_time'}, inplace=True)
    st.subheader('Stats per agent')
    st.dataframe(agents.style.highlight_max(axis=0, subset=['Calls', 'Avg_Talking_Time', 'Avg_Waiting_Time', 'Avg_Total_Time', 'Total_Talking_Time', 'Total_Waiting_Time', 'Total_Time'], color='#e4b090'))
    
    st.subheader('Calls per Agent excluding internal calls')
    all_info_excl_internal = all_call_info[all_call_info['type_x'] != 'internal']
    agents_ext = all_info_excl_internal.groupby(['agent_id', 'agent_name']).agg(Calls=('id', 'count'), Avg_Talking_Time=('talking_time', 'mean'), Avg_Waiting_Time=('waiting_time', 'mean'), Avg_Total_Time=('total_time', 'mean'),
                                                                                Total_Talking_Time=('talking_time', 'sum'), Total_Waiting_Time=('waiting_time', 'sum'), Total_Time=('total_time', 'sum')).reset_index()
    # Crosstab to count 'missed' and 'answered' calls
    status_counts = pd.crosstab(index=[all_info_excl_internal['agent_id'], all_info_excl_internal['agent_name']], columns=all_info_excl_internal['status'])

    # Merge the two DataFrames
    agents_ext = pd.merge(agents_ext, status_counts, how='left', left_on=['agent_id', 'agent_name'], right_index=True)

    # If you want to rename the status columns for clarity
    agents_ext.rename(columns={'answered': 'Answered_Calls', 'missed': 'Missed_Calls'}, inplace=True)
    #agents_ext.rename(columns={'id': 'calls', 'talking_time': 'avg_talking_time', 'waiting_time': 'avg_waiting_time', 'total_time': 'avg_total_time'}, inplace=True)
    st.dataframe(agents_ext.style.highlight_max(axis=0, subset=['Calls', 'Avg_Talking_Time', 'Avg_Waiting_Time', 'Avg_Total_Time', 'Total_Talking_Time', 'Total_Waiting_Time', 'Total_Time'], color='#e4b090'))
    
    st.subheader("Unanswered unique number calls per agent (except internal)")
    unanswered_calls = all_call_info[all_call_info['status'] == 'missed']
    unanswered_call_counts = unanswered_calls.groupby(['agent_id', 'agent_name']).agg(Unanswered_Cdr_UniqCount=('id', 'nunique')).reset_index()
    st.dataframe(unanswered_call_counts.style.highlight_max(axis=0, subset=['Unanswered_Cdr_UniqCount'], color='yellow'))
    
    st.subheader("Distribution of unanswered calls per ring time (0-10s, 10-20s, etc.)")
    unanswered_calls_ring = unanswered_calls[['id', 'ringing_time']]
    # input missing values with 0
    unanswered_calls_ring['ringing_time'] = unanswered_calls_ring['ringing_time'].fillna(0)

    # Define your bins (adjust according to your data range and needs)
    bins = [-1, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, float('inf')]
    labels = ['0-10', '10-20', '20-30', '30-40', '40-50', '50-60', '60-70', '70-80', '80-90', '90-100', '100+']

    # Categorize ringing time into bins
    unanswered_calls_ring['ringing_time_pct'] = pd.cut(unanswered_calls_ring['ringing_time'], bins, labels=labels)
    # streamlit ploting (barplot)
    st.bar_chart(unanswered_calls_ring['ringing_time_pct'].value_counts(), color="#92644d")
    
    st.subheader("Distribution of unanswered calls per hour and agent")
    all_call_info[['answered_at', 'ended_at']] = all_call_info[['answered_at', 'ended_at']].apply(pd.to_datetime, errors='coerce')
    all_call_info['answered_at'] = all_call_info['answered_at'].dt.tz_localize(None)
    all_call_info['ended_at'] = all_call_info['ended_at'].dt.tz_localize(None)

    # Distribution of unanswered calls per hour and agent
    # Create a new column with the hour of the day
    all_call_info['hour'] = all_call_info['answered_at'].dt.hour

    # Show distribution of unanswered calls per hour and agent in a pivot table
    unanswered_calls_hour = all_call_info[all_call_info['status'] == 'missed'].pivot_table(index=['agent_id', 'agent_name'], columns='hour', values='id', aggfunc='count', margins=True, margins_name='Total').fillna(0).astype(int)
    st.dataframe(unanswered_calls_hour, width=1000, height=500)
    unanswered_calls_hour.reset_index(inplace=True)
    unanswered_calls_hour.set_index(["agent_id", "agent_name"], inplace=True)
    unanswered_calls_hour.drop(index='Total', inplace=True)
    unanswered_calls_hour.drop(columns='Total', inplace=True)
    # Number of agents
    # Group data by Agent
    long_format = unanswered_calls_hour.reset_index().melt(id_vars=['agent_id', 'agent_name'], var_name='Hour', value_name='Values')
    grouped = long_format.groupby('agent_name')

    # Determine the grid size
    num_agents = len(grouped)
    nrows = int(num_agents**0.5)
    ncols = int(num_agents / nrows) + (num_agents % nrows > 0)

    # Iterate and display charts
    for i in range(nrows):
        cols = st.columns(ncols)
        for j in range(ncols):
            index = i*ncols + j
            if index < num_agents:
                agent_name = list(grouped.groups.keys())[index]
                agent_data = grouped.get_group(agent_name)
                chart = create_altair_chart(agent_name, agent_data)
                cols[j].altair_chart(chart, use_container_width=True)

    unanswered_calls_hour_t = unanswered_calls_hour.T
    unanswered_calls_hour_t['Total'] = unanswered_calls_hour_t.sum(axis=1)

    chart = alt.Chart(unanswered_calls_hour_t.reset_index()).mark_bar(
        color="#045e68",
        size=40,  # Equivalent to bar width
        opacity=0.8,
        cornerRadiusTopLeft=3,
        cornerRadiusTopRight=3
    ).encode(
        x=alt.X('hour:N', axis=alt.Axis(title='Hour', labelFontSize=17, titleFontSize=17)),
        y=alt.Y('Total:Q', axis=alt.Axis(title='Count', labelFontSize=17, titleFontSize=17)),
        tooltip=['hour', 'Total']
    ).properties(
        title="Distribution of Unanswered Calls per Hour",
        width=700,  # Adjust as needed
        height=400,  # Adjust as needed

    ).configure_title(
        fontSize=20
    )

    # Display the chart in Streamlit
    st.altair_chart(chart, use_container_width=True)
                
def main():
    # set the title of the window
    st.set_page_config(page_title='CloudTalk - Mitsis Group Analytics',layout='wide')
    st.title('CloudTalk - Mitsis Group Analytics')
    # two columns
    col1, col2 = st.columns(2)
    start_date = col1.date_input('Start date', value=dt.datetime.now(), key='start_date')
    start_time = col1.time_input('Start time', value=dt.time(0, 0, 0), key='start_time')
    end_date = col2.date_input('End date', key='end_date')
    end_time = col2.time_input('End time', value=dt.time(23, 59, 59), key='end_time')
    st.write('Click the button below to call the API')

    if st.button('Call API'):
        # make start and end dates into datetimes
        start_date = dt.datetime.combine(start_date, start_time)
        end_date = dt.datetime.combine(end_date, end_time)
        call_api(start_date, end_date)

if __name__ == '__main__':
    main()
