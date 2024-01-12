import pandas as pd
import os
import requests
import dotenv
import streamlit as st
dotenv.load_dotenv()

@st.cache_data
def make_api_call(url, type_url="history", **kwargs):
    """_summary_

    Args:
        url (_type_): _description_
        type_url (str, optional): _description_. Defaults to "history". Available options: history, call, agents
    """
    idd = os.environ.get('API_USER')
    pw = os.environ.get('API_PW')
    date_from = kwargs.get('date_from')
    date_to = kwargs.get('date_to')
    page = kwargs.get('page', 1)
    limit = kwargs.get('limit', 1000)
    # for call specific info
    callId = kwargs.get('call_id')
    if type_url == "call":
        r = requests.get(url.format(callId=callId), auth=(idd, pw))
        if r.status_code != 200:
            print("Error: " + str(r.status_code))
        return r.status_code, r
    elif type_url == 'agents':
        r = requests.get(url, auth=(idd, pw))
        if r.status_code != 200:
            print("Error: " + str(r.status_code))
        return r.status_code, r


def edit_data(answered_df, missed_df):
    answered_df.drop(columns=['Contact', 'Ratings', 'BillingCall', 'Notes', 'Tags', 'Agent','CallNumber'], inplace=True)
    missed_df.drop(columns=['Contact', 'Ratings', 'BillingCall', 'Notes', 'Tags', 'Agent','CallNumber'], inplace=True)
    answered_df = pd.json_normalize(answered_df['Cdr'])    
    missed_df = pd.json_normalize(missed_df['Cdr'])
    
    answered_df['status'] = 'answered'
    missed_df['status'] = 'missed'
    
    df = pd.concat([answered_df, missed_df])
    unique_cdr_ids = df['id'].unique()
    
    call_info = pd.DataFrame()
    api_url = "https://analytics-api.cloudtalk.io/api/calls/{callId}"
    st.markdown("""
        <style>
        .stProgress .st-am {
            background-color: #c68b66;
        }
        </style>
        """, unsafe_allow_html=True)
    progress_bar = st.progress(0, text=f"Fetching call info...({len(unique_cdr_ids)} calls)")
    false_requests = 0
    for idx, call_id in enumerate(unique_cdr_ids):
        # add a progress bar
        progress_bar.progress(idx/len(unique_cdr_ids), text=f"Fetching call info...({idx}/{len(unique_cdr_ids)} calls)")
        status, r = make_api_call(api_url, type_url="call", call_id=call_id)
        if status != 200:
            false_requests += 1
            continue
        info = r.json()['call_times'] 
        info['direction'] = r.json()['direction']
        info['cdr_id'] = r.json()['cdr_id']
        info['type'] = r.json()['type']
        info['status'] = r.json()['status']
        info['location'] = r.json()['internal_number']['name']
        call_info = pd.concat([call_info, pd.DataFrame([info])], ignore_index=True)
    
    progress_bar.progress(1.0, text="Done!")
    st.markdown(f"""
                ### API Call stasts
                * Error call requests: **{false_requests} out of {len(unique_cdr_ids)}**
                * Total Answered calls retrieved: **{len(answered_df)}**
                * Total Missed calls retrieved: **{len(missed_df)}**
                """)
    
    # Data cleaning
    df['id'] = df['id'].astype(str)
    call_info['cdr_id'] = call_info['cdr_id'].astype(str)

    all_call_info = pd.merge(df, call_info, left_on='id', right_on='cdr_id', how='left')
    
    # agents 
    status, r = make_api_call("https://my.cloudtalk.io/api/agents/index.json", type_url="agents")
    agents = pd.DataFrame(r.json()['responseData']['data'])
    agents = pd.json_normalize(agents['Agent'])
    agents = agents[['id', 'firstname', 'lastname', 'email']]
    agents.rename(columns={'id': 'agent_id', 'firstname':'agent_name', 'lastname':'agent_site'}, inplace=True)
    all_call_info = pd.merge(all_call_info, agents, left_on='user_id', right_on='agent_id', how='left')
    all_call_info['agent_name'] = all_call_info['agent_name'].fillna(all_call_info['location'])
    all_call_info['agent_id'] = all_call_info['agent_id'].fillna('999')
    
    all_call_info['status_y'] = all_call_info['status_y'].fillna(all_call_info['status_x'])
    all_call_info['talking_time_y'] = all_call_info['talking_time_y'].fillna(all_call_info['talking_time_x'])
    all_call_info['waiting_time_y'] = all_call_info['waiting_time_y'].fillna(all_call_info['waiting_time_x'])
    all_call_info.drop(columns=['status_x'], inplace=True)
    all_call_info.drop(columns=['talking_time_x'], inplace=True)
    all_call_info.drop(columns=['waiting_time_x'], inplace=True)
    all_call_info.rename(columns={'status_y': 'status', 'talking_time_y':'talking_time', 'waiting_time_y':'waiting_time'}, inplace=True)
    
    return all_call_info
