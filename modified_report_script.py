from docx.shared import Inches
from docx import Document
import matplotlib.pyplot as plt
import os
import requests
import pandas as pd
import warnings
warnings.filterwarnings("ignore")
pd.set_option("display.max_columns", None)
pd.set_option("display.max_rows", None)
pd.options.display.float_format = "{:,.2f}".format
import seaborn as sns
from matplotlib.ticker import MaxNLocator
import matplotlib.pyplot as plt
sns.set_style("darkgrid")



# Creating a new Word document
doc = Document()


def ETL(date_from, date_to):
    api_url = "https://my.cloudtalk.io/api/calls/index.json"

    # credentials
    idd = "X4THUIGLZTFDECBZKV6NMSD"
    pw = "GgOQV5tsSpjKPm3n;Ikczv6aW1ZTfJBH!q78M2XRuL"

    page = 1
    limit = 1000

    answered_df = pd.DataFrame()
    missed_df = pd.DataFrame()

    while True:
        print("Page: " + str(page))

        # make requests

        answered = requests.get(
            api_url,
            auth=(idd, pw),
            params={
                "page": page,
                "limit": limit,
                "date_from": date_from,
                "date_to": date_to,
                "status": "answered",
            },
        )

        page_count = answered.json()["responseData"]["pageCount"]

        # print(r.json())

        if answered.status_code != 200:
            print("Error: " + str(answered.status_code))
            break

        # convert to dataframe
        answered_df = pd.concat(
            [answered_df, pd.DataFrame(answered.json()["responseData"]["data"])]
        )
        page += 1
        if page >= page_count:
            break
    page = 1

    while True:
        print("Page: " + str(page))
        # make requests
        missed = requests.get(
            api_url,
            auth=(idd, pw),
            params={
                "page": page,
                "limit": limit,
                "date_from": date_from,
                "date_to": date_to,
                "status": "missed",
            },
        )
        page_count = missed.json()["responseData"]["pageCount"]
        if missed.status_code != 200:
            print("Error: " + str(missed.status_code))
            break
        # convert to dataframe
        missed_df = pd.concat(
            [missed_df, pd.DataFrame(missed.json()["responseData"]["data"])]
        )
        page += 1
        if page >= page_count:
            break

    answered_df.drop(
        columns=[
            "Contact",
            "Ratings",
            "BillingCall",
            "Notes",
            "Tags",
            "Agent",
            "CallNumber",
        ],
        inplace=True,
    )

    missed_df.drop(
        columns=[
            "Contact",
            "Ratings",
            "BillingCall",
            "Notes",
            "Tags",
            "Agent",
            "CallNumber",
        ],
        inplace=True,
    )

    answered_df = pd.json_normalize(answered_df["Cdr"])
    missed_df = pd.json_normalize(missed_df["Cdr"])


    answered_df["status"] = "answered"
    missed_df["status"] = "missed"

    df = pd.concat([answered_df, missed_df])

    unique_cdr_ids = df["id"].unique()

    # Call API to get call details
    call_info = pd.DataFrame()
    error_calls = 0
    api_url = "https://analytics-api.cloudtalk.io/api/calls/{callId}"
    for idx, call_id in enumerate(unique_cdr_ids):
        print(f"{idx}/{len(unique_cdr_ids)}")
        r = requests.get(api_url.format(callId=call_id), auth=(idd, pw))
        if r.status_code != 200:
            print("Error: " + str(r.status_code))
            error_calls += 1
            continue

        info = r.json()["call_times"]
        info["direction"] = r.json()["direction"]
        info["cdr_id"] = r.json()["cdr_id"]
        info["type"] = r.json()["type"]
        info["status"] = r.json()["status"]
        info["location"] = r.json()["internal_number"]["name"]
        call_info = pd.concat([call_info, pd.DataFrame([info])], ignore_index=True)

    df["id"] = df["id"].astype(str)
    call_info["cdr_id"] = call_info["cdr_id"].astype(str)
    all_call_info = pd.merge(df, call_info, left_on="id", right_on="cdr_id", how="left")
    
    # Agents API

    api_url = "https://my.cloudtalk.io/api/agents/index.json"
    r = requests.get(api_url, auth=(idd, pw))
    agents = pd.DataFrame(r.json()["responseData"]["data"])
    agents = pd.json_normalize(agents["Agent"])
    agents = agents[["id", "firstname", "lastname", "email"]]
    agents.rename(
        columns={"id": "agent_id", "firstname": "agent_name", "lastname": "agent_site"},
        inplace=True,
    )
    all_call_info = pd.merge(
        all_call_info, agents, left_on="user_id", right_on="agent_id", how="left"
    )
    all_call_info["agent_name"] = all_call_info["agent_name"].fillna(
        all_call_info["location"]
    )
    all_call_info["agent_id"] = all_call_info["agent_id"].fillna("999")
    all_call_info["status_y"] = all_call_info["status_y"].fillna(all_call_info["status_x"])
    all_call_info["talking_time_y"] = all_call_info["talking_time_y"].fillna(
        all_call_info["talking_time_x"]
    )
    all_call_info["waiting_time_y"] = all_call_info["waiting_time_y"].fillna(
        all_call_info["waiting_time_x"]
    )
    all_call_info.drop(columns=["status_x"], inplace=True)
    all_call_info.drop(columns=["talking_time_x"], inplace=True)
    all_call_info.drop(columns=["waiting_time_x"], inplace=True)

    all_call_info.rename(
        columns={
            "status_y": "status",
            "talking_time_y": "talking_time",
            "waiting_time_y": "waiting_time",
        },
        inplace=True,
    )


    all_call_info["talking_time"] = all_call_info["talking_time"].astype(float)
    all_call_info["waiting_time"] = all_call_info["waiting_time"].astype(float)
    all_call_info["total_time"] = all_call_info["total_time"].astype(float)

    agents = (
        all_call_info.groupby(["agent_id", "agent_name"])
        .agg(
            Calls=("id", "count"),
            Avg_Talking_Time=("talking_time", "mean"),
            Avg_Waiting_Time=("waiting_time", "mean"),
            Avg_Total_Time=("total_time", "mean"),
            Total_Talking_Time=("talking_time", "sum"),
            Total_Waiting_Time=("waiting_time", "sum"),
            Total_Time=("total_time", "sum"),
        )
        .reset_index()
    )

    # Crosstab to count 'missed' and 'answered' calls

    status_counts = pd.crosstab(
        index=[all_call_info["agent_id"], all_call_info["agent_name"]],
        columns=all_call_info["status"],
    )


    # Merge the two DataFrames

    agents = pd.merge(
        agents,
        status_counts,
        how="left",
        left_on=["agent_id", "agent_name"],
        right_index=True,
    )


    # If you want to rename the status columns for clarity

    agents.rename(
        columns={"answered": "Answered_Calls", "missed": "Missed_Calls"}, inplace=True
    )

    # 1: Stats per agent

    all_info_excl_internal = all_call_info[all_call_info["type_x"] != "internal"]

    agents_ext = (
        all_info_excl_internal.groupby(["agent_id", "agent_name"])
        .agg(
            Calls=("id", "count"),
            Avg_Talking_Time=("talking_time", "mean"),
            Avg_Waiting_Time=("waiting_time", "mean"),
            Avg_Total_Time=("total_time", "mean"),
            Total_Talking_Time=("talking_time", "sum"),
            Total_Waiting_Time=("waiting_time", "sum"),
            Total_Time=("total_time", "sum"),
        )
        .reset_index()
    )

    # Crosstab to count 'missed' and 'answered' calls

    status_counts = pd.crosstab(
        index=[all_info_excl_internal["agent_id"], all_info_excl_internal["agent_name"]],
        columns=all_info_excl_internal["status"],
    )


    # Merge the two DataFrames

    agents_ext = pd.merge(
        agents_ext,
        status_counts,
        how="left",
        left_on=["agent_id", "agent_name"],
        right_index=True,
    )


    # If you want to rename the status columns for clarity

    agents_ext.rename(
        columns={"answered": "Answered_Calls", "missed": "Missed_Calls"}, inplace=True
    )

    # 2: Unanswered calls per agent excluding internal calls

    unanswered_calls = all_call_info[all_call_info["status"] == "missed"]
    unanswered_call_counts = (
        unanswered_calls.groupby(["agent_id", "agent_name"])
        .agg(Unanswered_UniqCount=("id", "nunique"))
        .reset_index()
    )

    # 3: Unanswered unique number calls per agent (no internal)
    unanswered_calls_ring = unanswered_calls[["id", "ringing_time"]]
    # input missing values with 0
    unanswered_calls_ring["ringing_time"] = unanswered_calls_ring["ringing_time"].fillna(0)
    # Define your bins (adjust according to your data range and needs)
    bins = [-1, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, float("inf")]

    labels = [
        "0-10",
        "10-20",
        "20-30",
        "30-40",
        "40-50",
        "50-60",
        "60-70",
        "70-80",
        "80-90",
        "90-100",
        "100+",
    ]


    # Categorize ringing time into bins

    unanswered_calls_ring["ringing_time_pct"] = pd.cut(
        unanswered_calls_ring["ringing_time"], bins, labels=labels
    )


    # plot the distribution

    # 4: Distribution of unanswered calls per ring time (0-10s, 10-20s, etc.)
    fig = (
        unanswered_calls_ring["ringing_time_pct"]
        .value_counts()
        .sort_index()
        .plot.bar(
            color="#045e68",
            figsize=(20, 6),
            title="Distribution of Ringing Time",
            rot=25,
            fontsize=17,
            xlabel="Ringing Time (seconds)",
            ylabel="Count",
            width=0.8,
            edgecolor="#00141b",
            linewidth=1.2,
            zorder=3,
            grid=True,
            alpha=0.8,
        )
    )
    # make the title bigger
    fig.title.set_size(20)
    # shorten the margins betwwn the fig and the plot
    #fig.figure.subplots_adjust(bottom=0.2, top=0.2, left=0.2, right=0.2)
    fig.figure.tight_layout()
    fig.figure.savefig("unanswered_calls_ring_plot.png")


    all_call_info["answered_at"] = pd.to_datetime(all_call_info["answered_at"], errors="coerce")
    all_call_info["ended_at"] = pd.to_datetime(all_call_info["ended_at"], errors="coerce")
    all_call_info["answered_at"] = all_call_info["answered_at"].dt.tz_localize(None)
    all_call_info["ended_at"] = all_call_info["ended_at"].dt.tz_localize(None)


    # Distribution of unanswered calls per hour and agent
    # Create a new column with the hour of the day
    all_call_info["hour"] = all_call_info["answered_at"].dt.hour

    # Show distribution of unanswered calls per hour and agent in a pivot table
    unanswered_calls_hour = (
        all_call_info[all_call_info["status"] == "missed"]
        .pivot_table(
            index=["agent_id", "agent_name"],
            columns="hour",
            values="id",
            aggfunc="count",
            margins=True,
            margins_name="Total",
        )
        .fillna(0)
        .astype(int)
    )

    unanswered_calls_hour.reset_index(inplace=True)
    unanswered_calls_hour.set_index(["agent_id", "agent_name"], inplace=True)
    unanswered_calls_hour.drop(index="Total", inplace=True)
    unanswered_calls_hour.drop(columns="Total", inplace=True)

    # 5: Distibution of unanswered calls per hour and agent
    unanswered_calls_hour
    num_agents = len(unanswered_calls_hour)

    # Creating a grid for subplots
    ncols = 4
    nrows = int(num_agents / ncols) + (num_agents % ncols > 0)


    # Creating subplots
    fig, axes = plt.subplots(nrows, ncols, figsize=(20, 15), constrained_layout=True)

    # Flattening the axes array for easy iteration
    axes = axes.flatten()

    # Define the custom bar color
    bar_color = "#c68b66"

    # Plotting in each subplot
    for i, (index, row) in enumerate(unanswered_calls_hour.iterrows()):
        axes[i].bar(row.index.astype(str), row.values, color=bar_color)  # Convert index to string and set bar color
        axes[i].set_title(f"{index[1]}", fontsize=18)

        # Set one ylabel per row
        if i % ncols == 0:
            # Make the values int
            axes[i].set_ylabel("Count", fontsize=16)
            
        axes[i].set_xlabel("Hour")
        axes[i].tick_params(labelsize=12)

    # Hiding any unused subplots
    for j in range(i + 1, len(axes)):
        axes[j].axis("off")
        # make the yaxis integer only
        axes[j].yaxis.set_major_locator(MaxNLocator(integer=True))

    fig.tight_layout()
    # 6:  Distribution of unanswered calls per hour and agent
    plt.savefig("hours.png")

    # plot the distribution of hours
    # 7: Distribution of unanswered calls per hour
    # transpose the dataframe
    unanswered_calls_hour_t = unanswered_calls_hour.T
    unanswered_calls_hour_t['Total'] = unanswered_calls_hour_t.sum(axis=1)
    
    fig1 = plt.figure(figsize=(20, 6))
    fig1 = (unanswered_calls_hour_t['Total'].plot.bar(
            color="#045e68",
            figsize=(20, 6),
            title="Distribution of Unanswered Calls per Hour",
            rot=0,
            fontsize=17,
            xlabel="Hour",
            ylabel="Count",
            width=0.8,
            edgecolor="#00141b",
            linewidth=1.2,
            zorder=3,
            grid=True,
            alpha=0.8,
        ))
    
    # make the title bigger
    fig1.title.set_size(20)
    fig1.figure.tight_layout()
    fig1.figure.savefig("unanswered_calls_hour_plot.png")

    return answered_df, missed_df, error_calls, agents, agents_ext, unanswered_call_counts, unanswered_calls_ring, unanswered_calls_hour