# %%
# Start counting timer
import time
start_time = time.time()

# %%
# Import Library
import pandas as pd
import numpy as np
import os
import plotly.graph_objects as go

# %%
# Define Constant Variable
## Column label for 3-stage submission
col_name = ['ProjectName', 'doc_no', 'doc_title', 'weight', 'rev1_plan', 'rev1_actual', 'rev1_return', 'rev2_plan', 'rev2_actual', 'rev2_return', 'rev3_plan', 'rev3_actual', 'Discipline']
## Variable of date
today = pd.to_datetime('today')
## Label for each project
label_project1 = 'PROJECT 1'
label_project2 = 'PROJECT 2'
label_project3 = 'PROJECT 3'
label_project4 = 'PROJECT 4'
## This Python script absolut directory address
def absolut_path():
    try:
        script_dir = os.path.dirname(os.path.realpath(__file__)) # This script works cross platform except for Jupyter Notebook
        return script_dir
    except:
        script_dir = os.path.dirname(os.path.realpath('__file__')) # Works with Jupyter Notebook
        return script_dir

# %%
# Function to Create Each Project Dataframe
def project_query(file_path, sheet_name, col_no, top_rows, bottom_rows, project_name):
    # Add Excel project schedule relative address    
    abs_file_path = os.path.join(absolut_path(), file_path)
    # Read sheet in Excel file with only required columns
    df_raw = pd.read_excel(abs_file_path, sheet_name=sheet_name, usecols=col_no)
    # Drop all empty rows then drop unused top and bottom rows
    df = df_raw.dropna(axis=0, how='all').iloc[top_rows:bottom_rows, :].copy()
    # Insert column 3rd stage for 2-stage submission with blank value
    if len(df.columns) == 10:
        df.insert(loc=len(df.columns)-1, column='rev3_plan', value=np.nan)
        df.insert(loc=len(df.columns)-1, column='rev3_actual', value=np.nan)
    # Insert ProjectName column at the start of dataframe table
    df.insert(loc=0, column='ProjectName', value=project_name)
    # Rename table column as per predefined list
    df.columns = col_name
    # Filter row that contain particular word(s)
    df = df[~df['doc_title'].str.contains('delete', case=False, na=False, regex=False)]
    # Change 'weight' column data type to np.float64
    df.loc[:, 'weight'] = df.loc[:, 'weight'].astype(np.float64)
    # Change date columns data type to datetype
    df.iloc[:, 4:-1] = df.iloc[:, 4:-1].fillna(np.NaN).apply(pd.to_datetime, errors='coerce')
    # Strip string from deliverable name
    df.loc[:, 'doc_title'] = df.loc[:, 'doc_title'].str.strip()
    # Reset row index
    df.reset_index(drop=True, inplace=True)
    # Return function to dataframe variable
    return df

# %%
# Apply the function to each Excel project schedule
# Project 1
df_project1 = project_query('excel/project1_schedule.xlsx', 'MDR', [1,2,4,6,7,11,14,15,19,22], 5, -7, label_project1)

# Project 2
df_project2 = project_query('excel/project2_schedule.xlsx', 'MDR', [1,2,5,7,8,12,15,16,20,23], 5, -4, label_project2)

# Project 3
df_project3 = project_query('excel/project3_schedule.xlsx', 'MDR', [1,2,7,9,10,14,17,18,22,25], 5, -1, label_project3)

# Project 4
df_project4 = project_query('excel/project4_schedule.xlsx', 'MDR', [1,2,4,9,12,14,15,18,20,23,26,27], 3, -2, label_project4)

# %%
# Join (append / concat) all tables
# List of projects dataframe to be joined
df_list = [df_project1, df_project2, df_project3, df_project4]
# Join all projects dataframe
df_all = pd.concat(df_list, ignore_index=True)

# %%
# Show list of late deliverables
df_late_list = df_all[
    ((df_all['rev1_plan'] < today) & (df_all['rev1_actual'].isnull() & df_all['rev1_return'].isnull())) |
    ((df_all['rev2_plan'] < today) & (df_all['rev1_actual'].notnull() | df_all['rev1_return'].notnull()) & (df_all['rev2_actual'].isnull() & df_all['rev2_return'].isnull())) |
    ((df_all['rev3_plan'] < today) & (df_all['rev1_actual'].notnull() | df_all['rev1_return'].notnull()) & (df_all['rev2_actual'].notnull() | df_all['rev2_return'].notnull()) & (df_all['rev3_actual'].isnull()))
].drop(['weight', 'rev1_return', 'rev2_return'], axis=1)
df_late_list

# %%
# Quantity of late deliverables each revision per project per discipline
df_late_count_r1 = df_all.loc[
    ((df_all['rev1_plan'] < today) & (df_all['rev1_actual'].isnull() & df_all['rev1_return'].isnull())), ['ProjectName', 'Discipline', 'rev1_plan']]
df_late_count_r2 = df_all.loc[
    ((df_all['rev2_plan'] < today) & (df_all['rev1_actual'].notnull() | df_all['rev1_return'].notnull()) & (df_all['rev2_actual'].isnull() & df_all['rev2_return'].isnull())), ['ProjectName', 'Discipline', 'rev2_plan']]
df_late_count_r3 = df_all.loc[
    ((df_all['rev3_plan'] < today) & (df_all['rev1_actual'].notnull() | df_all['rev1_return'].notnull()) & (df_all['rev2_actual'].notnull() | df_all['rev2_return'].notnull()) & (df_all['rev3_actual'].isnull())), ['ProjectName', 'Discipline', 'rev3_plan']]
# %%
df_late_count_r1 = df_late_count_r1.groupby(['ProjectName', 'Discipline']).count()
df_late_count_r2 = df_late_count_r2.groupby(['ProjectName', 'Discipline']).count()
df_late_count_r3 = df_late_count_r3.groupby(['ProjectName', 'Discipline']).count()
# %%
# Merge all late deliverable count table
df_late_count_list = [df_late_count_r1, df_late_count_r2, df_late_count_r3]
df_late_count = pd.concat(df_late_count_list, join='outer', axis=1, sort=False).reset_index(drop=False)
# %%
# Convert rev column value type to int
df_late_count[['rev1_plan', 'rev2_plan', 'rev3_plan']] = (
    df_late_count[['rev1_plan', 'rev2_plan', 'rev3_plan']].fillna(0)
    .astype(int)
)

# %% 
# Create bar chart for late deliverables quantity
# Define the data
late_bar1 = go.Bar(x=[df_late_count['ProjectName'], df_late_count['Discipline']], y=df_late_count['rev1_plan'], name='Rev. 1', text=df_late_count['rev1_plan'], textposition='auto')
late_bar2 = go.Bar(x=[df_late_count['ProjectName'], df_late_count['Discipline']], y=df_late_count['rev2_plan'], name='Rev. 2', text=df_late_count['rev2_plan'], textposition='auto')
late_bar3 = go.Bar(x=[df_late_count['ProjectName'], df_late_count['Discipline']], y=df_late_count['rev3_plan'], name='Rev. 3', text=df_late_count['rev3_plan'], textposition='auto')
# Combine the data into list
data_late = [late_bar1, late_bar2, late_bar3]
# Set the layout
layout_late = go.Layout(title='Late Deliverable Quantity',
                        xaxis_title='Projects',
                        yaxis_title='Quantity',
                        barmode='stack')
# Combine Data & Layout into Chart
fig_late = go.Figure(data = data_late, layout = layout_late)
# Show the chart
fig_late.show()

# %%
# Project Progress (Planned vs. Actual)
# Create a Dataframe Copy from df_all
df_all_prog = df_all.copy()
# Planned Progress
# Zero progress
df_all_prog.loc[(df_all['rev1_plan'] > today), 'rev1_pln_prog'] = 0
df_all_prog.loc[(df_all['rev2_plan'] > today), 'rev2_pln_prog'] = 0
df_all_prog.loc[(df_all['rev3_plan'].isnull() | (df_all['rev3_plan'] > today)), 'rev3_pln_prog'] = 0
# 2-stage submission project
df_all_prog.loc[(df_all['rev3_plan'].isnull() & (df_all['rev1_plan'] <= today)), 'rev1_pln_prog'] = df_all.weight * 0.75
df_all_prog.loc[(df_all['rev3_plan'].isnull() & (df_all['rev2_plan'] <= today)), 'rev2_pln_prog'] = df_all.weight * 0.25
# 3-stage submission project
df_all_prog.loc[(df_all['rev3_plan'].notnull() & (df_all['rev1_plan'] <= today)), 'rev1_pln_prog'] = df_all.weight * 0.4
df_all_prog.loc[(df_all['rev3_plan'].notnull() & (df_all['rev2_plan'] <= today)), 'rev2_pln_prog'] = df_all.weight * 0.3
df_all_prog.loc[(df_all['rev3_plan'].notnull() & (df_all['rev3_plan'] <= today)), 'rev3_pln_prog'] = df_all.weight * 0.3
# Sum of Weight Column of Planned Progress
df_all_prog.loc[:, 'sum_pln_prog'] = df_all_prog.rev1_pln_prog + df_all_prog.rev2_pln_prog + df_all_prog.rev3_pln_prog
# Actual progress
# Zero progress
df_all_prog.loc[(df_all['rev1_actual'].isnull() & df_all['rev1_return'].isnull()), 'rev1_act_prog'] = 0
df_all_prog.loc[(df_all['rev2_actual'].isnull() & df_all['rev2_return'].isnull()), 'rev2_act_prog'] = 0
df_all_prog.loc[df_all['rev3_actual'].isnull(), 'rev3_act_prog'] = 0
# 2-stage submission project
df_all_prog.loc[(df_all['rev3_plan'].isnull() & (df_all['rev1_actual'].notnull() | df_all['rev1_return'].notnull())), 'rev1_act_prog'] = df_all.weight * 0.75
df_all_prog.loc[(df_all['rev3_plan'].isnull() & (df_all['rev2_actual'].notnull() | df_all['rev2_return'].notnull())), 'rev2_act_prog'] = df_all.weight * 0.25
# 3-stage submission project
df_all_prog.loc[(df_all['rev3_plan'].notnull() & (df_all['rev1_actual'].notnull() | df_all['rev1_return'].notnull())), 'rev1_act_prog'] = df_all.weight * 0.4
df_all_prog.loc[(df_all['rev3_plan'].notnull() & (df_all['rev2_actual'].notnull() | df_all['rev2_return'].notnull())), 'rev2_act_prog'] = df_all.weight * 0.3
df_all_prog.loc[(df_all['rev3_plan'].notnull() & df_all['rev3_actual'].notnull()), 'rev3_act_prog'] = df_all.weight * 0.3
# Sum of Weight Column of Actual Progress
df_all_prog.loc[:, 'sum_act_prog'] = df_all_prog.rev1_act_prog + df_all_prog.rev2_act_prog + df_all_prog.rev3_act_prog
# Create GroupBy table
df_all_prog = df_all_prog.groupby(['ProjectName'])[['sum_pln_prog', 'sum_act_prog']].sum()
# Reset index
df_all_prog = df_all_prog.reset_index(drop=False)

# %%
# Create bar chart for Planned vs. Actual Progress
# Format the label value of each bar in chart into percentage
lbl_pln_prog = pd.Series(['{0:.2f}%'.format(val * 100) for val in df_all_prog.sum_pln_prog], index = df_all_prog.index)
lbl_act_prog = pd.Series(['{0:.2f}%'.format(val * 100) for val in df_all_prog.sum_act_prog], index = df_all_prog.index)
# Define the data
pln_prog_bar = go.Bar(x=df_all_prog['ProjectName'], y=df_all_prog['sum_pln_prog'], name='Planned', text=lbl_pln_prog, textposition='outside')
act_prog_bar = go.Bar(x=df_all_prog['ProjectName'], y=df_all_prog['sum_act_prog'], name='Actual', text=lbl_act_prog, textposition='outside')
# Combine the data into list
data_prog = [pln_prog_bar, act_prog_bar]
# Set the layout
layout_prog = go.Layout(
    title='Planned vs. Actual Progress',
    xaxis_title='Projects',
    yaxis=dict(
        title='Progress',
        tickformat='.2%',
        range=[0, 1.1],
    ),
)
# Combine Data & Layout into Chart
fig_prog = go.Figure(data = data_prog, layout = layout_prog)
# Show the chart
fig_prog.show()

# %%
# Function to create S-Curve Dataframe
def project_scurve(start_date, end_date, df_project):
    # Add variable for start and finish time to be included in date_range variable
    str_df = pd.to_datetime(start_date)
    end_df = pd.to_datetime(end_date)
    dt_df = pd.Series(pd.date_range(start=str_df, end=end_df, freq='7D'))
    # Add blank list for planned and actual progress time series
    pln_df = []
    act_df = []    
    for i in dt_df.index:            
        if df_project['rev3_plan'].count() > 0:
            ## 3 revisions
            # Planned progress  
            pln_df.append(df_project
                          .loc[((df_project['rev1_plan'] <= dt_df[i]) 
                                & df_project['rev1_plan'].notnull() 
                                & df_project['rev3_plan'].notnull()), 'weight']
                          .sum()*0.4 
                          + df_project
                          .loc[((df_project['rev2_plan'] <= dt_df[i]) 
                                & df_project['rev2_plan'].notnull() 
                                & df_project['rev3_plan'].notnull()), 'weight']
                          .sum()*0.3 
                          + df_project
                          .loc[((df_project['rev3_plan'] <= dt_df[i]) 
                                & df_project['rev3_plan'].notnull()), 'weight']
                          .sum()*0.3
                          )
            # Actual progress
            act_df.append(df_project
                          .loc[((df_project['rev1_actual'] <= dt_df[i]) 
                                & (df_project['rev1_actual'].notnull() | df_project['rev1_return'].notnull()) 
                                & df_project['rev3_plan'].notnull() 
                                & (dt_df[i] <= pd.to_datetime('today'))), 'weight']
                          .sum()*0.4 
                          + df_project
                          .loc[((df_project['rev2_actual'] <= dt_df[i]) 
                                & (df_project['rev2_actual'].notnull() | df_project['rev2_return'].notnull()) 
                                & df_project['rev3_plan'].notnull() 
                                & (dt_df[i] <= pd.to_datetime('today'))), 'weight']
                          .sum()*0.3 
                          + df_project
                          .loc[((df_project['rev3_actual'] <= dt_df[i]) 
                                & df_project['rev3_actual'].notnull() 
                                & df_project['rev3_plan'].notnull() 
                                & (dt_df[i] <= pd.to_datetime('today'))), 'weight']
                          .sum()*0.3
                          )
        else:
            ## 2 revisions
            # Planned progress  
            pln_df.append(df_project
                          .loc[((df_project['rev1_plan'] <= dt_df[i]) 
                                & df_project['rev1_plan'].notnull() 
                                & df_project['rev3_plan'].isnull()), 'weight']
                          .sum()*0.75 
                          + df_project
                          .loc[((df_project['rev2_plan'] <= dt_df[i]) 
                                & df_project['rev2_plan'].notnull() 
                                & df_project['rev3_plan'].isnull()), 'weight']
                          .sum()*0.25
                          )        
            # Actual progress
            act_df.append(df_project
                          .loc[((df_project['rev1_actual'] <= dt_df[i]) 
                                & (df_project['rev1_actual'].notnull() | df_project['rev1_return'].notnull()) 
                                & df_project['rev3_plan'].isnull() 
                                & (dt_df[i] <= pd.to_datetime('today'))), 'weight']
                          .sum()*0.75 
                          + df_project
                          .loc[((df_project['rev2_actual'] <= dt_df[i]) 
                                & (df_project['rev2_actual'].notnull() | df_project['rev2_return'].notnull()) 
                                & df_project['rev3_plan'].isnull() 
                                & (dt_df[i] <= pd.to_datetime('today'))), 'weight']
                          .sum()*0.25
                          )
    dt_df = pd.Series(dt_df, name='Date')
    pln_df = pd.Series(pln_df, name='Planned')
    act_df = pd.Series(act_df, name='Actual')    
    df = pd.concat([dt_df, pln_df, act_df], axis=1)
    df.loc[:,'Actual'] = df.loc[:,'Actual'].replace(0, np.nan)
    return df

# %%
# Define max date comparing max planned finish date and current date
def max_date(col):
    if col.max() > today:
        return col.max().strftime('%d-%b-%y')
    else:
        return 'today'

# %%
# Run the S-Curve Dataframe Function for Each Project
scurve_project1 = project_scurve(df_project1['rev1_plan'].min().strftime('%d-%b-%y'), max_date(df_project1['rev2_plan']), df_project1)
scurve_project2 = project_scurve(df_project2['rev1_plan'].min().strftime('%d-%b-%y'), max_date(df_project2['rev2_plan']), df_project2)
scurve_project3 = project_scurve(df_project3['rev1_plan'].min().strftime('%d-%b-%y'), max_date(df_project3['rev2_plan']), df_project3)
scurve_project4 = project_scurve(df_project4['rev1_plan'].min().strftime('%d-%b-%y'), max_date(df_project4['rev3_plan']), df_project4)

# %%
# Function to convert S-Curve DataFrame to Plotly Line Chart
def scurve_project(scurve_df, project_name):
    # Format the label value of each point in chart into percentage
    lbl_pln_sc = pd.Series(['{0:.2f}%'.format(val * 100) for val in scurve_df['Planned']], index = scurve_df.index)
    lbl_act_sc = pd.Series(['{0:.2f}%'.format(val * 100) for val in scurve_df['Actual']], index = scurve_df.index)
    # Define the data
    df_pln_sc = go.Scatter(x=scurve_df['Date'], y=scurve_df['Planned'], name='Planned', text=lbl_pln_sc, textposition='top center', mode="lines+markers+text",)
    df_act_sc = go.Scatter(x=scurve_df['Date'], y=scurve_df['Actual'], name='Actual', text=lbl_act_sc, textposition='top center', mode="lines+markers+text",)
    # Combine the data into list
    data_sc = [df_pln_sc, df_act_sc]
    # Set the layout
    layout_sc = go.Layout(
        title='{} S-Curve'.format(project_name),
        xaxis_title='Date',
        yaxis=dict(
            title='Progress',
            tickformat='.2%',
            range=[0, 1.1],
        ),
    )
    # Combine Data & Layout into Chart
    sc_project = go.Figure(data = data_sc, layout = layout_sc)
    # Return the chart
    return sc_project
    
# %%
# Run the S-Curve Conversion Function for Each Project
sc_project1 = scurve_project(scurve_project1, label_project1)
sc_project1.show()
sc_project2 = scurve_project(scurve_project2, label_project2)
sc_project2.show()
sc_project3 = scurve_project(scurve_project3, label_project3)
sc_project3.show()
sc_project4 = scurve_project(scurve_project4, label_project4)
sc_project4.show()

# %%
print('Script finished running in:')
print('--- %s seconds ---' % (time.time() - start_time))

# %%
