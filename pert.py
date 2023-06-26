import math
from collections import defaultdict
import networkx as nx
import xlrd

projects = defaultdict(list)
tasks = {}  # format = activity | duration
edges = []  # edges for network diagram
nodes = []  # nodes for network diagram


# Critical Path Method Forward Pass Function
# EF -> Earliest Finish
# ES -> Earliest Start
def forward_pass(mydata):
    es = {}
    ef = {}
    for act, prd in projects.items():  # for every activity in the project
        # print("act: ", act)
        # print("predecessors: ", prd)
        dur = tasks[act]  # Duration time
        temp_values = []  # holds temporary values
        # print("prd: ", prd)
        for p in prd:
            if p == 'Start' or p == 'NONE':
                es[act] = 0.0
                ef[act] = es[act] + dur  # Earliest Start Plus Duration
            else:
                temp_values.append(ef[p])  # holds all the EF values of predecessors
                es[act] = max(temp_values)  # Earliest Start = Highest EF value of predecessors
                ef[act] = es[act] + dur  # Earliest Start Plus Duration
        temp_values = []  # reset temporary values

    # Update dataFrame:
    mydata['ES'] = es
    mydata['EF'] = ef


# Critical Path Method Backward Pass function
# LS -> Latest Start
# LF -> Latest Finish
def backward_pass(mydata, successors):
    ls = {}
    lf = {}
    for act, succ in reversed(successors.items()):  # get successors for each activity
        # print("act: ", act)
        # print("successors: ", succ)
        dur = tasks[act]  # Duration time
        # print("dur: ", dur)
        completion_time = tasks['Finish']
        temp_values = []  # holds temporary values

        if act == 'Start':  # skip Start Node
            lf[act] = 0.0  # set to zero
            ls[act] = 0.0
            continue

        for s in reversed(succ):
            if s == 'Finish':
                lf[act] = completion_time  # set end nodes with completion time
                ls[act] = round((lf[act] - dur), 2)  # Latest Finish minus Duration
            else:
                # print("activity: ", act)
                # print("LS list of successors: ", succ)
                temp_values.append(ls[s])  # holds all the LS values of successors
                lf[act] = min(temp_values)  # Latest Finish = Lowest LS of the successors
                ls[act] = round((lf[act] - dur), 2)  # Latest Start minus Duration
        temp_values = []  # reset temporary values

    # Update dataFrame:
    mydata['LS'] = ls
    mydata['LF'] = lf


def get_completion_time(ef):
    return max(ef.values())


# G => graph
def get_successors(G):
    successors = defaultdict(list)
    g_nodes = G.nodes()
    for n in g_nodes:
        if G.successors(n) is not None:
            for s in list(G.successors(n)):
                successors[n].append(s)

    return successors


# compute for SLACK values and CRITICAL state
# Slack = LS - ES or LF - ES of activity
# Critical activity when slack value is zero
def compute_slack_values(mydata):
    # compute slack value for each activity/task
    slack = {}
    critical = {}
    critical_path = {}
    for act in tasks:
        if act == 'Start' or act == 'Finish':  # skip Start and Finish Nodes
            slack[act] = 0.0
            continue
        slack[act] = round((mydata['LF'][act] - mydata['EF'][act]), 2)
        if slack[act] == 0:  # activity is critical if slack value is equal to zero
            critical[act] = "YES"
            critical_path[act] = act
        else:
            critical[act] = "NO"

    mydata['SLACK'] = slack
    mydata['CRITICAL'] = critical
    mydata['CRITICAL_PATH'] = critical_path


# checks if string is a valid number
def is_number(string):
    try:
        string = str(string)
        if string.isnumeric():
            return True
        elif float(string):
            return True
        else:
            return False
    except ValueError:
        return False


# draws the Project Network
# G => graph
def draw_network(G):
    G.add_nodes_from(nodes)
    G.add_edges_from(edges)

    end_nodes = [n for n, o in G.out_degree() if o == 0]
    tasks['Start'] = 0
    tasks['Finish'] = 0
    for n in end_nodes:
        G.add_edge(n, 'Finish')
        projects['Finish'].append(n)

    nx.set_node_attributes(G, {'Finish': {"color": "orange"}})
    nx.set_node_attributes(G, {'Start': {"color": "orange"}})

    node_colors = [c for c in nx.get_node_attributes(G, "color").values()]
    colors = [i for c, i in nx.get_edge_attributes(G, "color").items()]
    pos = nx.spring_layout(G, k=10 / math.sqrt(G.order()))
    nx.draw_networkx_nodes(G, pos, node_size=1000, node_color=node_colors)
    nx.draw_networkx_edges(G, pos, edgelist=G.edges(), edge_color=colors, arrowsize=30)
    nx.draw_networkx_labels(G, pos)


def read_data_file(filename):
    book = xlrd.open_workbook(filename)  # open excel data
    sheet = book.sheet_by_index(0)  # get sheet from excel

    for row in range(1, sheet.nrows):  # skips header row
        data = sheet.row_slice(row)
        # data[0] = activity
        # data[1] = predecessor
        # data[2] = duration
        act = ""
        pre = []
        # activity value per row
        if is_number(data[0].value):
            act = str(int(data[0].value))
        else:
            act = data[0].value

        # predecessor value per row
        if is_number(data[1].value):
            pre.append(str(int(data[1].value)))
        else:
            pre = data[1].value.split(',')

        # duration value per row
        dur = data[2].value
        tasks[act] = float(dur)
        nodes.append((act, {"color": "lightblue"}))

        # activity can have multiple predecessors
        for p in pre:
            if p == "NONE":
                p = "Start"
            edges.append((p, act, {"color": "k"}))  # k is black color
            projects[act].append(p)


# draws critical parth of the network
# set red nodes and edges for critical path
# G => graph
def draw_critical_path(G, mydata):
    critical_path = list(mydata['CRITICAL_PATH'].values())
    g_nodes = G.nodes()
    # set red color for critical nodes
    for e in g_nodes:
        if e in critical_path:
            nx.set_node_attributes(G, {e: {"color": "red"}})
        else:
            nx.set_node_attributes(G, {e: {"color": "lightblue"}})

    nx.set_node_attributes(G, {'Start': {"color": "red"}})
    nx.set_node_attributes(G, {'Finish': {"color": "red"}})

    node_colors = [c for c in nx.get_node_attributes(G, "color").values()]
    g_edges = G.edges()
    # set red color for critical edges
    critical_edges = get_critical_edges(critical_path)
    for e in g_edges:
        if e in critical_edges:
            nx.set_edge_attributes(G, {e: {"color": "red"}})
        else:
            nx.set_edge_attributes(G, {e: {"color": "k"}})

    colors = [i for c, i in nx.get_edge_attributes(G, "color").items()]
    pos = nx.spring_layout(G, k=10 / math.sqrt(G.order()))
    nx.draw_networkx_nodes(G, pos, node_size=1000, node_color=node_colors)
    nx.draw_networkx_edges(G, pos, edgelist=G.edges(), edge_color=colors, arrowsize=30)
    nx.draw_networkx_labels(G, pos)


# gets critical edges
def get_critical_edges(critical_path):
    critical_edges = [('Start', critical_path[0])]
    length = len(critical_path)
    # Start edge
    for i in range(length - 1):
        critical_edges.append((critical_path[i], critical_path[i + 1]))

    # Finish edge
    critical_edges.append((critical_path[length - 1], 'Finish'))

    return critical_edges


# removes Start and Finish values
def remove_start_finish_data(mydata):
    data = {}
    for k, v in mydata.items():
        if 'Start' in v:
            del v['Start']
        if 'Finish' in v:  # remove Finish and Start
            del v['Finish']
        if k == 'CRITICAL_PATH':  # skips CRITICAL_PATH values
            continue
        sorted(v)
        data[k] = v
    return data


def format_critical_path(path):
    arrow = "->"
    return arrow.join(path)


def get_project_data():
    projects_data = {}
    for p, v in projects.items():
        projects_data[p] = [p, v, tasks[p]]
    return projects_data
