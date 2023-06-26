# prerequisites
# pip install networkx, matplotlib, xlrd and pandas

import matplotlib.pyplot as plt
import pandas as pd
from pert import *

# read project's data saved i ProjectData.xls file
file = "ProjectData.xls"
read_data_file(file)

projects_data = get_project_data()
proj_df = pd.DataFrame.from_dict(projects_data, columns=['Activity', 'Predecessor', 'Duration'], orient="index")
# Show projects data in excel
print(proj_df.to_string(index=False))

# set tasks/activities data
mydata = defaultdict(list)
mydata['TASKS'] = {k: k for k in tasks.keys()}

G = nx.DiGraph()
plt.figure("Figure 1: Project Network")
draw_network(G)

# Computes the Earliest Start Time and Earliest Finish Time
forward_pass(mydata)

# Get the Project Completion Time
completion_time = get_completion_time(mydata['EF'])

# Set the Finish Time
tasks['Finish'] = completion_time
# print("Finish time: ", tasks['Finish'])

# Compute the Earliest Start Time and Earliest Finish Time
successors = get_successors(G)
backward_pass(mydata, successors)

# Compute Slack and Get Critical Nodes/Edges
compute_slack_values(mydata)

# print("nodes_colors: ", node_colors)
# print("nodes: ", nodes)
# print("edges: ", edges)
# print("successors: ", successors)
# print("COMPLETION TIME: ", completion_time)
# print("ES: ", mydata['ES'])
# print("EF: ", mydata['EF'])
# print("LS: ", mydata['LS'])
# print("LF: ", mydata['LF'])
# print("SLACK: ", mydata['SLACK'])
# print("CRITICAL: ", mydata['CRITICAL'])
# print("CRITICAL_PATH: ", mydata['CRITICAL_PATH'])
G2 = G.copy()

plt.figure("Figure 2: Critical Path")  # Figure 2 with Critical Path
draw_critical_path(G2, mydata)

columns = {
    'CODE': ['TASKS', 'ES', 'EF', 'LS', 'LF', 'SLACK', 'CRITICAL'],
    'NAME': ['Activity', 'ES', 'EF', 'LS', 'LF', 'Slack Value', 'Is Critical']
}

data = remove_start_finish_data(mydata)
df = pd.DataFrame(data, columns=columns['CODE'])
# Show results
print(df.to_string(index=False))

print("\nCompletion Time: ", completion_time)
print("\nCritical Path: ", format_critical_path(mydata['CRITICAL_PATH']))

# Show figures
plt.show()

