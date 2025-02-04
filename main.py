import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta

def calculate_statistics(df, start_date, end_date, accessory_name):
    if start_date is None or end_date is None:
        start_date, end_date = df["Date"].min(), df["Date"].max()

    df = df[(df["Date"] >= start_date) & (df["Date"] <= end_date)].copy()
    df["Date_Only"] = df["Date"].dt.date  # Extract date for grouping

    open_sessions = 0
    session_durations = []
    daily_open_times = {}
    open_time = None

    for i, row in df.iterrows():
        if row["State"] == 1:  # Contact Open
            if open_time is not None:
                time_gap = (row["Date"] - open_time).total_seconds() / 3600  # Convert to hours
                if time_gap > 24:
                    open_time = None  # Reset due to potential measurement error
                    continue
            if open_time is None:
                open_time = row["Date"]
                open_sessions += 1
        elif row["State"] == 0 and open_time is not None:  # Contact Closed
            duration = (row["Date"] - open_time).total_seconds() / 60
            if duration <= 1440:  # Ensure duration is realistic (less than 24 hours)
                session_durations.append(duration)
                day = row["Date_Only"]
                if day not in daily_open_times:
                    daily_open_times[day] = 0
                daily_open_times[day] += duration
            open_time = None

    df_daily = pd.DataFrame(list(daily_open_times.items()), columns=["Date_Only", "Open_Duration"])
    df_daily["Closed_Duration"] = 1440 - df_daily["Open_Duration"]

    mean_open_sessions = open_sessions / df_daily.shape[0] if df_daily.shape[0] > 0 else 0
    mean_open_minutes = df_daily["Open_Duration"].mean()
    mean_closed_minutes = df_daily["Closed_Duration"].mean()
    mean_open_time_per_session = sum(session_durations) / len(session_durations) if session_durations else 0

    return {
        "Accessory Name": accessory_name,
        "From": start_date.strftime("%Y-%m-%d %H:%M:%S"),
        "To": end_date.strftime("%Y-%m-%d %H:%M:%S"),
        "Number of Days": (end_date - start_date).days + 1,
        "Mean Minutes Open Per Day": mean_open_minutes,
        "Mean Minutes Closed Per Day": mean_closed_minutes,
        "Mean Opening Sessions Per Day": mean_open_sessions,
        "Mean Open Time Per Session": mean_open_time_per_session
    }

def plot_contact_data(files, combined_plot, start_date, end_date, show_stats):
    all_stats = []

    if combined_plot:
        plt.figure(figsize=(12, 5))

    for file in files:
        df, accessory_name = read_contact_file(file)

        if start_date is not None and end_date is not None:
            df = df[(df["Date"] >= start_date) & (df["Date"] <= end_date)]

        if combined_plot:
            plt.step(df["Date"], df["State"], where="post", label=accessory_name, linewidth=2)
        else:
            plt.figure(figsize=(12, 5))
            plt.step(df["Date"], df["State"], where="post", label=accessory_name, linewidth=2)
            plt.xlabel("Date/Time")
            plt.ylabel("Contact State (1 = Open, 0 = Closed)")
            plt.title(f"Contact Open/Closed State Over Time - {accessory_name}")
            plt.yticks([0, 1], labels=["Closed", "Open"])
            plt.xticks(rotation=45)
            plt.grid(True, linestyle="--", alpha=0.6)
            plt.legend()
            st.pyplot(plt)

        if show_stats:
            stats = calculate_statistics(df, start_date, end_date, accessory_name)
            all_stats.append(stats)
            st.write(f"### Statistics for {accessory_name}")
            st.write(stats)

    if combined_plot:
        plt.xlabel("Date/Time")
        plt.ylabel("Contact State (1 = Open, 0 = Closed)")
        plt.title("Contact Open/Closed State Over Time - Combined")
        plt.yticks([0, 1], labels=["Closed", "Open"])
        plt.xticks(rotation=45)
        plt.grid(True, linestyle="--", alpha=0.6)
        plt.legend()
        st.pyplot(plt)

    if len(all_stats) > 1:
        overall_stats = pd.DataFrame(all_stats).drop(
            columns=["Accessory Name", "From", "To", "Number of Days"]).mean().to_dict()
        st.write("### Overall Statistics (Mean Over All Uploaded Files)")
        st.write(overall_stats)

def read_contact_file(file):
    metadata_df = pd.read_excel(file, sheet_name="Contact", nrows=3, header=None)  # Read first 3 metadata lines
    accessory_name = metadata_df.iloc[0, 0].split(": ")[1] if len(metadata_df.columns) > 1 else "Unknown"

    df = pd.read_excel(file, sheet_name="Contact", skiprows=3)  # Skip metadata rows
    df["Date"] = pd.to_datetime(df["Date"], format="%Y-%m-%dT%H:%M:%S")
    df = df.sort_values("Date")
    df["State"] = df["Contact"].map({"Open": 1, "Closed": 0})
    return df, accessory_name

def read_netatmo_file(file):
    metadata_df = pd.read_excel(file, sheet_name="Worksheet", nrows=2, header=None)
    accessory_name = metadata_df.iloc[1, 3].split(" ")[1]

    df = pd.read_excel(file, sheet_name="Worksheet", skiprows=2)
    df["Date"] = pd.to_datetime(df["Timezone : Europe/Berlin"], format="%Y/%m/%d %H:%M:%S")
    df = df.sort_values("Date")
    return df, accessory_name


def add_fake_states(df, start_date, end_date):
    if start_date is None:
        return df

    start = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    end = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)

    if df.empty:
        return df

    first_state = df.iloc[0]["State"]
    last_state = df.iloc[-1]["State"]

    fake_start = pd.DataFrame([{"Date": start, "State": last_state}])
    fake_end = pd.DataFrame([{"Date": end, "State": last_state}])

    df = pd.concat([fake_start, df, fake_end], ignore_index=True)
    return df

def filter_data_by_time(df, start_date, end_date):
    if start_date is None or end_date is None:
        return df  # Return unfiltered if no time range is set
    return df[(df["Date"] >= start_date) & (df["Date"] <= end_date)]


def plot_multiple_data(files, selected_measurements, start_date, end_date, show_stats):
    all_stats = []

    fig, ax1 = plt.subplots(figsize=(12, 5))
    ax2 = ax1.twinx()

    for file, measurement in selected_measurements:
        df, name = (read_contact_file(file) if "Contact" in file.name else read_netatmo_file(file))
        df = filter_data_by_time(df, start_date, end_date)
        df = add_fake_states(df, start_date, end_date) if "Contact" in file.name else df

        if measurement in df.columns:
            if "Contact" in file.name:
                ax2.step(df["Date"], df[measurement], where="post", label=f"{name} - {measurement}", linewidth=2,
                         linestyle='--')
            else:
                ax1.plot(df["Date"], df[measurement], label=f"{name} - {measurement}", linewidth=2, color="purple")

        if "Contact" in file.name and show_stats:
            stats = calculate_statistics(df, start_date, end_date, name)
            all_stats.append(stats)
            st.write(f"### Statistics for {name}")
            st.write(stats)

    ax1.set_xlabel("Date/Time")
    ax1.set_ylabel("Netatmo Values")
    ax2.set_ylabel("Contact Sensor State (0/1)")

    ax1.legend(loc="upper left")
    ax2.legend(loc="upper right")

    plt.title("Combined Sensor Data")
    plt.grid(True, linestyle="--", alpha=0.6)
    st.pyplot(fig)

    if len(all_stats) > 1:
        overall_stats = pd.DataFrame(all_stats).drop(
            columns=["Accessory Name", "From", "To", "Number of Days"]).mean().to_dict()
        st.write("### Overall Statistics (Mean Over All Uploaded Files)")
        st.write(overall_stats)


def main():
    st.title("State and Sensor Data Plotter")

    uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx"], accept_multiple_files=True)
    time_filter = st.selectbox("Select Time Filter", ["All", "Today", "Last Week", "Last Month", "Last Year", "Custom"])

    custom_start, custom_end = None, None
    if time_filter == "Custom":
        custom_start = st.date_input("Start Date", datetime.now() - timedelta(days=7))
        custom_end = st.date_input("End Date", datetime.now())

    netatmo_and_eve = False
    for file in uploaded_files:
        if "Netatmo" in file.name:
            netatmo_and_eve = True

    combined_plot = True
    if not netatmo_and_eve:
        combined_plot = st.checkbox("Show all files in one plot", value=True)

    show_stats = st.checkbox("Show Statistics", value=True)

    now = datetime.now()
    if time_filter == "Last Week":
        start_date = now - timedelta(weeks=1)
        end_date = now
    elif time_filter == "Last Month":
        start_date = now - timedelta(days=30)
        end_date = now
    elif time_filter == "Last Year":
        start_date = now - timedelta(days=365)
        end_date = now
    elif time_filter == "Today":
        start_date = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end_date = now.replace(hour=23, minute=59, second=59, microsecond=999999)
    elif time_filter == "Custom":
        if custom_start and custom_end:
            start_date = datetime.combine(custom_start, datetime.min.time())
            end_date = datetime.combine(custom_end, datetime.max.time())
        else:
            start_date, end_date = None, None
    else:  # "All"
        start_date, end_date = None, None

    if netatmo_and_eve:
        selected_measurements = []
        if uploaded_files:
            for file in uploaded_files:
                df, name = (read_contact_file(file) if "Contact" in file.name else read_netatmo_file(file))
                if "Contact" in file.name:
                    measurement = "State"
                else:
                    measurement = st.selectbox(f"Select measurement for {name}", df.columns[2:], key=file.name)
                selected_measurements.append((file, measurement))

        if uploaded_files:
            plot_multiple_data(uploaded_files, selected_measurements, start_date, end_date, show_stats)
    else:
        if uploaded_files:
            if len(uploaded_files) > 6:
                st.error("Please upload a maximum of 6 files.")
            else:
                plot_contact_data(uploaded_files, combined_plot, start_date, end_date, show_stats)

if __name__ == "__main__":
    main()
