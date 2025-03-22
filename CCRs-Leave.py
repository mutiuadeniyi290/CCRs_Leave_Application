import pandas as pd
import numpy as np
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
import os

# Sample leave application data
datafile = "employee_leave_2025.xlsx"

# Check if the file exists
if not os.path.exists(datafile):
    st.error("Error: The file 'employee_leave_2025.xlsx' was not found.")
    st.stop()  # Stop execution if file is missing

# Convert to DataFrame
leave_df = pd.read_excel(datafile)

# Ensure required columns exist
required_columns = {'Employee_Name', 'Start_Date', 'End_Date', 'Reason'}
if not required_columns.issubset(leave_df.columns):
    st.error(f"Error: Missing required columns: {required_columns - set(leave_df.columns)}")
    st.stop()

# Convert dates to datetime
leave_df['Start_Date'] = pd.to_datetime(leave_df['Start_Date'])
leave_df['End_Date'] = pd.to_datetime(leave_df['End_Date'])

# Calculate working days (Monday-Friday)
leave_df['Working Days'] = leave_df.apply(lambda row: np.busday_count(row['Start_Date'].date(), row['End_Date'].date()), axis=1)


# Calculate Leave Duration in Days
leave_df['Leave_Duration'] = leave_df.apply(
    lambda row: np.busday_count(row['Start_Date'].date(), row['End_Date'].date()) + 1,
    axis=1)

# Extract Month Name from Start_Date
leave_df['Month_Name'] = leave_df['Start_Date'].dt.strftime('%B')

# Define month order for sorting
month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
               'July', 'August', 'September', 'October', 'November', 'December']

# Convert Month_Name column to categorical
leave_df['Month_Name'] = pd.Categorical(leave_df['Month_Name'], categories=month_order, ordered=True)

# Sidebar Filters
selected_employee = st.sidebar.selectbox("Select Employee", ["All"] + sorted(leave_df["Employee_Name"].dropna().unique()))
selected_leave_types = st.sidebar.multiselect("Select Leave Type", leave_df["Reason"].dropna().unique(), default=leave_df["Reason"].dropna().unique())

# Apply filters
filtered_df = leave_df[
    ((leave_df['Employee_Name'] == selected_employee) | (selected_employee == "All")) &
    (leave_df['Reason'].isin(selected_leave_types))
]

# Streamlit app layout
st.title("📊 Customer Service Leave Trend Analysis by Month & Type")

st.write("""
This dashboard visualizes the number of leave applications per month, categorized by leave type.
It helps HR track trends and manage leave planning.
""")

# Display filtered data table
with st.expander("📋 View Filtered Leave Data"):
    st.dataframe(filtered_df)

# Calculate key leave metrics
total_leaves = len(filtered_df)
total_leave_days = filtered_df['Leave_Duration'].sum() if total_leaves > 0 else 0
average_leave_days = round(filtered_df['Leave_Duration'].mean(), 2) if total_leaves > 0 else 0

# Layout for displaying metrics
col1, col2, col3 = st.columns(3)
col1.metric(label="📊 Total Leave Applications", value=total_leaves)
col2.metric(label="📅 Total Leave Days", value=total_leave_days)
col3.metric(label="📊 Avg Leave Duration (Days)", value=average_leave_days)

st.subheader("Leave summary")
leave_status_summary = filtered_df['Reason'].value_counts()


total_leave = {
    "Annual": leave_status_summary.get("Annual Leave", 0),
    "Sick": leave_status_summary.get("Sick Leave", 0),
     "Study": leave_status_summary.get("Study Leave", 0),
    "Casual": leave_status_summary.get("Casual Leave", 0),
    "Maternity": leave_status_summary.get("Maternity Leave", 0),
    "Paternity": leave_status_summary.get("Paternity Leave", 0)
}

st.subheader("Leave Summary by Type")
st.write(f"**Total Annual Leave**: {total_leave['Annual']}")
st.write(f"**Total Sick Leave**: {total_leave['Sick']}") 
st.write(f"**Total Study Leave**: {total_leave['Study']}") 
st.write(f"**Total Casual Leave**: {total_leave['Casual']}")
st.write(f"**Total Maternity Leave**: {total_leave['Maternity']}")
st.write(f"**Total Paternity Leave**: {total_leave['Paternity']}")




# --- Leave Trend by Month ---
leave_trend = filtered_df.groupby(['Month_Name', 'Reason']).size().reset_index(name='Leave_Count')


if not leave_trend.empty:
    leave_pivot = leave_trend.pivot(index='Month_Name', columns='Reason', values='Leave_Count').fillna(0)

    st.subheader("📈 Leave Trend by Month & Type")
    fig, ax = plt.subplots(figsize=(10, 5))
    leave_pivot.plot(kind='bar', stacked=True, colormap='Set2', ax=ax)

    ax.set_xlabel("Month")
    ax.set_ylabel("Number of Leave Applications")
    ax.set_title(f"Leave Trend by Month & Type ({selected_employee})")
    ax.legend(title="Reason")
    ax.set_xticklabels(ax.get_xticklabels(), rotation=45)
    ax.grid(axis='y', linestyle='--', alpha=0.7)
    st.pyplot(fig)

    # --- Heatmap for Monthly Leave Trends ---
    st.subheader("🔥 Monthly Leave Trend Heatmap")
    heatmap_data = filtered_df.groupby(['Month_Name', 'Reason'])['Leave_Duration'].sum().unstack().fillna(0)

    fig, ax = plt.subplots(figsize=(10, 5))
    sns.heatmap(heatmap_data, cmap="Blues", annot=True, fmt=".1f", linewidths=0.5, ax=ax)
    ax.set_title("Total Leave Days by Month & Type")
    st.pyplot(fig)
else:
    st.warning("⚠️ No data available for the selected filters.")



# Define the correct month order
month_order = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Ensure 'Month_Name' is a categorical variable with the correct order
leave_trend = filtered_df.groupby(['Employee_Name', 'Month_Name']).size().reset_index(name='Leave_Taken_Each_Month_By_Staff')

# Group by Month to get total leave requests per month
monthly_leave = leave_trend.groupby('Month_Name')['Leave_Taken_Each_Month_By_Staff'].sum().reset_index()

# Convert 'Month_Name' to categorical and sort accordingly
monthly_leave['Month_Name'] = pd.Categorical(monthly_leave['Month_Name'], categories=month_order, ordered=True)
monthly_leave = monthly_leave.sort_values('Month_Name')

# Display results
st.subheader('Determine the month with the highest number of leave requests')
st.table(monthly_leave)



# ---- Leave Type Breakdown (Pie Chart) ----
st.subheader("📊 Leave Type Breakdown")
if not leave_status_summary.empty:

    fig, ax = plt.subplots()
    colors = sns.color_palette("Set2", len(leave_status_summary))
    pd.Series(leave_status_summary).plot(kind='pie', autopct='%1.1f%%', startangle=90, colors=colors, ax=ax)
    ax.set_ylabel("")
    ax.set_title("Percentage of Leave Types")
    st.pyplot(fig)

else:
    st.warning("⚠️ No leave data available for pie chart.")    

# ---- Leave Distribution by Employee ----
if not filtered_df.empty:
    leave_distribution = filtered_df.groupby(['Reason', 'Employee_Name'])['Leave_Duration'].sum().reset_index().sort_values('Leave_Duration')
    
    st.subheader("📊 Leave Distribution by Type & Employee")
    # st.write(leave_distribution)

    with st.expander("📋 View Leave Breakdown by Employee"):
        st.write(leave_distribution)
else:
    st.warning("⚠️ No data available for leave distribution.")


leave_counts = leave_df['Employee_Name'].value_counts()

# Get the employee with the highest leave count
top_employee = leave_counts.idxmax()  # Employee with most leaves
top_leave_count = leave_counts.max()  # Number of leaves taken

st.subheader("🏆 Employee with Most Leave")
st.write(f"**{top_employee}** has taken the most leave, with **{top_leave_count}** applications.")

# Display full leave count table
with st.expander("📋 View Leave Counts for All Employees"):
    st.write(leave_counts)


# Count the number of leaves taken by each employee
leave_counts = leave_df['Employee_Name'].value_counts()

# Get the employee with the least leave count
least_employee = leave_counts.idxmin()  # Employee with least leaves
least_leave_count = leave_counts.min()  # Minimum number of leaves taken

st.subheader("🥇 Employee with Least Leave")
st.write(f"**{least_employee}** has taken the least leave, with **{least_leave_count}** application(s).")

# Display full leave count table
with st.expander("📋 View Leave Counts for All Employees"):
    st.write(leave_counts)    


 # Calculate total leave days for each employee
leave_df['Leave_Duration'] = (leave_df['End_Date'] - leave_df['Start_Date']).dt.days + 1
leave_summary = leave_df.groupby('Employee_Name')['Leave_Duration'].sum().reset_index()

# Get the top 5 employees with the most leave taken
top_5_most_leave = leave_summary.nlargest(5, 'Leave_Duration')

# Get the bottom 5 employees with the least leave taken
bottom_5_least_leave = leave_summary.nsmallest(5, 'Leave_Duration')


st.markdown(
    "<hr style='border: 2px solid green;'>", unsafe_allow_html=True
)


# Display in Streamlit
st.subheader("🏆 Employees with Most & Least Leave Taken")

col1, col2 = st.columns(2)

with col1:
    st.write("**Top 5 Employees with Most Leave Taken**")
    st.dataframe(top_5_most_leave)

with col2:
    st.write("**Bottom 5 Employees with Least Leave Taken**")
    st.dataframe(bottom_5_least_leave)
   

