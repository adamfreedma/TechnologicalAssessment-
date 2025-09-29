"""
Yuli Tshuva
Populations based graphs.
"""

# Imports
import pandas as pd
import numpy as np
import os
from os.path import join
import networkx as nx
from pyvis.network import Network
import community

# Constants
ITAMAR = False
DATA_PATH = "switched_ids.xlsx"
FORMER_DATA_PATH = "mh_sems3.xlsx"
OUTPUT_DIR = "output"
SEQUENTIAL_NUMBER = "מספר שלי - ללא רווחים"
ID = "תעודת זהות - ללא רווחים"
NAME = "שם מלא"
CORRIDOR = "מסדרון"
CLASS = "מחלקה"
SEX = "מגדר"
SECTOR = "מגזר"
BACKGROUND = "רקע אקדמי"
TIRONUT = "מחלקה בטירונות"
CHOOSES = ["בחירה 1", "בחירה 2", "בחירה 3", "בחירה 4", "בחירה 5", "בחירה 6"]
POPULATIONS = [CORRIDOR, CLASS, SEX, SECTOR, BACKGROUND, TIRONUT]
COLORS = ["dodgerblue", "hotpink", "lightcoral", "turquoise", "springgreen", "cyan", "fuchsia", "orange", "blue",
          "gold", "lightpink", "lightblue", "red", "black"]
NODE_DISTANCE = 550
FILENAME = lambda x: join(OUTPUT_DIR, f"graph_of_{x}.html")

# Make the output dir
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Read the data
df = pd.read_excel(DATA_PATH)
df3 = pd.read_excel(FORMER_DATA_PATH)

df[SEQUENTIAL_NUMBER] = df[SEQUENTIAL_NUMBER].astype(int)
df3[SEQUENTIAL_NUMBER] = df3[SEQUENTIAL_NUMBER].astype(int)
for choose in CHOOSES:
    # Convert the column to int if possible
    df[choose] = df[choose].fillna(0).astype(int)
    df3[choose] = df3[choose].fillna(0).astype(int)

# For the great commandor Itamar
if ITAMAR:
    dct = {
        "זכר": "גבר",
        "נקבה": "אישה"
    }
    df["מגדר"] = df["מין"].apply(lambda x: dct[x])
    df.drop(columns=["מין"], inplace=True)
    df3["מגדר"] = df3["מין"].apply(lambda x: dct[x])
    df3.drop(columns=["מין"], inplace=True)

# Set dtype
df[SEQUENTIAL_NUMBER] = df[SEQUENTIAL_NUMBER].astype(int)
for choose in CHOOSES:
    # Convert the column to int if possible
    df[choose] = df[choose].fillna(0).astype(int)


def create_edges_list(df):
    # Set an edges list
    edges = []

    # Iterate over the rows of the dataframe
    for i in range(len(df)):
        # Set the head of the edge
        head = df[SEQUENTIAL_NUMBER].tolist()[i]
        # Iterate over the members it chose
        for j in range(6):
            # Find the j'th tail
            tail = df[CHOOSES[j]].tolist()[i]
            # Check if the tail is not nan
            if np.isnan(tail) or int(tail) == 0:
                continue
            # Make sure the head didn't wrote itself
            if head != tail:
                edges.append((head, tail))

    # Print the amount of edges found
    print(f"Edges found: {len(edges)}.")
    # Return the edges list
    return edges


# Get the list of edges
edges = create_edges_list(df)
former_edges = create_edges_list(df3)
# Create a map from sequential number to ID
id_map = dict(zip(df[SEQUENTIAL_NUMBER], df[ID]))
# Create a map from sequential number to name
name_map = dict(zip(df[SEQUENTIAL_NUMBER], df[NAME]))

# Set a dict of color maps
color_maps = {}

# Iterate over the populations
for population in POPULATIONS:
    # Get the unique values
    unique_values = df[population].unique()

    # Create a map from sequential number to population
    population_map = dict(zip(df[SEQUENTIAL_NUMBER], df[population]))

    # Create a map from value to color
    value_color_map = dict(zip(unique_values, COLORS[:len(unique_values)]))

    # Save the color map
    color_maps[population] = value_color_map

    # Create a graph for the population
    G = nx.DiGraph()
    G.add_edges_from(edges)

    # Compute in-degrees
    in_degrees = dict(G.in_degree())

    # Compute centrality
    centrality = nx.pagerank(G)
    max_cent = np.max(list(centrality.values()))
    centrality = {key: value / max_cent for key, value in centrality.items()}
    centralises = np.sort(np.array(list(centrality.values())))
    threshold = centralises[-8]

    # Create PyVis network
    net = Network(height="750px", width="100%", directed=True, notebook=False)
    net.from_nx(G)

    # Customize node size based on in-degree
    for node in net.nodes:
        id = node["id"]
        degree = in_degrees.get(node["id"], 0)
        node["size"] = 5 + 1.5 * degree  # Adjust scaling factor as needed
        node["title"] = f"Name: {name_map[id]}.\nIn-degree: {degree}.\nCentrality: {centrality[id]:.3f}."
        # Write inside the node its id
        node["label"] = f"{name_map[id]}"
        # The types with the label inside of it are:
        # ellipse, circle, database, box, text.
        # The ones with the label outside of it are:
        # image, circularImage, diamond, dot, star, triangle, triangleDown, square and icon.
        node["shape"] = "dot"
        node["color"] = value_color_map[population_map[id]]
        if centrality[id] >= threshold:
            node["shape"] = "diamond"

    # Increase spacing and control movement
    net.repulsion(
        node_distance=NODE_DISTANCE,
        central_gravity=0.02,
        spring_length=200,
        spring_strength=0.1,
        damping=0.09
    )

    # Show graph
    net.write_html(join(OUTPUT_DIR, f"graph_of_{population}.html"), notebook=False)

# Create a graph for the population
G = nx.DiGraph()
G.add_edges_from(edges)

# Compute in-degrees
in_degrees = dict(G.in_degree())

# Compute centrality
centrality = nx.pagerank(G)
max_cent = np.max(list(centrality.values()))
centrality = {key: value / max_cent for key, value in centrality.items()}
centralises = np.sort(np.array(list(centrality.values())))
threshold = centralises[-8]

# Create PyVis network
net = Network(height="750px", width="100%", directed=True, notebook=False)
net.from_nx(G)

# Customize node size based on in-degree
for node in net.nodes:
    id = node["id"]
    degree = in_degrees.get(node["id"], 0)
    node["size"] = 5 + 1.5 * degree  # Adjust scaling factor as needed
    node["title"] = f"Name: {name_map[id]}.\nIn-degree: {degree}.\nCentrality: {centrality[id]:.3f}."
    # Write inside the node its id
    node["label"] = f"{name_map[id]}"
    # The types with the label inside of it are:
    # ellipse, circle, database, box, text.
    # The ones with the label outside of it are:
    # image, circularImage, diamond, dot, star, triangle, triangleDown, square and icon.
    node["shape"] = "dot"
    node["color"] = "silver"
    if centrality[id] >= threshold:
        node["shape"] = "diamond"

# Increase spacing and control movement
net.repulsion(
    node_distance=NODE_DISTANCE,
    central_gravity=0.02,
    spring_length=200,
    spring_strength=0.1,
    damping=0.09
)

# Show graph
net.write_html(join(OUTPUT_DIR, "general_graph.html"), notebook=False)

# Write the color maps to a json file
with open(join(OUTPUT_DIR, "color_maps.txt"), "w", encoding="utf-8") as f:
    for pop in POPULATIONS:
        f.write(f"{pop}:\n")
        for key, value in color_maps[pop].items():
            f.write(f"\t{key}: {value}\n")
        f.write("\n")


def cluster_graph(edge_list):
    # Step 1: Create a directed graph
    G_directed = nx.DiGraph()
    G_directed.add_weighted_edges_from(edge_list)

    # Step 2: Convert to undirected for clustering
    G = G_directed.to_undirected()

    # Step 3: Apply Louvain community detection
    partition = community.best_partition(G, resolution=1.5)

    # Step 4: Group nodes by their community
    clusters = {}
    for node, comm_id in partition.items():
        clusters.setdefault(comm_id, []).append(node)

    return clusters, partition, G


# Create a graph for the population
G, former_G = nx.DiGraph(), nx.DiGraph()
G.add_edges_from(edges)
former_G.add_edges_from(former_edges)

# Compute in-degrees
in_degrees = dict(G.in_degree())
former_in_degrees = dict(former_G.in_degree())

# Compute centrality
centrality = nx.pagerank(G)
max_cent = np.max(list(centrality.values()))
centrality = {key: value / max_cent for key, value in centrality.items()}
centralises = np.sort(np.array(list(centrality.values())))
threshold = centralises[-8]

# Create PyVis network
net = Network(height="750px", width="100%", directed=True, notebook=False)
net.from_nx(G)

# Cluster graph
undirected_edges = {}
for edge in edges:
    if (edge[1], edge[0]) in undirected_edges:
        undirected_edges[(edge[1], edge[0])] = 2
    else:
        undirected_edges[edge] = 1
undirected_edges_list = [(edge[0], edge[1], undirected_edges[edge]) for edge in undirected_edges]
clusters, partition, _ = cluster_graph(undirected_edges_list)
value_color_map = dict(zip(list(set(partition.values())), COLORS[:len(clusters)]))

# Dictionary from id to (name, in-degree, centrality)
information_list = []

# Customize node size based on in-degree
for node in net.nodes:
    id = node["id"]
    degree = in_degrees.get(node["id"], 0)
    node["size"] = 5 + 1.5 * degree  # Adjust scaling factor as needed
    node["title"] = f"Name: {name_map[id]}.\nIn-degree: {degree}.\nCentrality: {centrality[id]:.3f}."
    information_list.append([name_map[id], degree, centrality[id]])
    # Write inside the node its id
    node["label"] = f"{name_map[id]}"
    # The types with the label inside of it are:
    # ellipse, circle, database, box, text.
    # The ones with the label outside of it are:
    # image, circularImage, diamond, dot, star, triangle, triangleDown, square and icon.
    node["shape"] = "dot"
    node["color"] = value_color_map[partition[id]]
    if centrality[id] >= threshold:
        node["shape"] = "diamond"

# Increase spacing and control movement
net.repulsion(
    node_distance=NODE_DISTANCE,
    central_gravity=0.02,
    spring_length=200,
    spring_strength=0.1,
    damping=0.09
)

# Save graph
net.write_html(join(OUTPUT_DIR, f"clustered_graph.html"), notebook=False)

df2 = pd.DataFrame(information_list, columns=["Name", "In-degree", "Centrality"])
df2.to_csv(join("output", "exported_data.csv"), encoding="utf-8-sig", index=False)

# Get the nodes
nodes = list(id_map.keys())

# Make a dir for personal graphs
os.makedirs(join(OUTPUT_DIR, "personal_graphs"), exist_ok=True)

for node1 in nodes:
    node_edges = [edge for edge in edges if node1 in edge]

    former_node_edges = [edge for edge in former_edges if ((node1 in edge) and (edge[0] in nodes) and (edge[1] in nodes))]
    merged_edges = node_edges + former_node_edges
    merged_edges = list(set(merged_edges))

    # Create a graph for the population
    G = nx.DiGraph()
    G.add_edges_from(merged_edges)

    # Create PyVis network
    net = Network(height="750px", width="100%", directed=True, notebook=False)
    net.from_nx(G)

    # Customize node size based on in-degree
    for node in net.nodes:
        id = node["id"]
        node["size"] = 5 + 1.5 * 6
        # Write inside the node its id
        node["label"] = f"{name_map[id]}"
        # The types with the label inside of it are:
        # ellipse, circle, database, box, text.
        # The ones with the label outside of it are:
        # image, circularImage, diamond, dot, star, triangle, triangleDown, square and icon.
        node["shape"] = "dot"
        node["color"] = "silver"

    for edge in net.edges:
        source = edge["from"]
        target = edge["to"]
        edge_exp = (source, target)

        # Example logic to change color based on source or target node
        check1, check2 = edge_exp in node_edges, edge_exp in former_node_edges
        if check1 and check2:
            edge["color"] = "purple"
        elif check1:
            edge["color"] = "green"
        else:
            edge["color"] = "red"

    # Increase spacing and control movement
    net.repulsion(
        node_distance=NODE_DISTANCE,
        central_gravity=0.02,
        spring_length=200,
        spring_strength=0.1,
        damping=0.09
    )

    # Show graph
    name_map = {k: v.replace("'", "") for k, v in name_map.items()}
    net.write_html(join(OUTPUT_DIR, "personal_graphs", f'{name_map[node1].replace(" ", "_")}.html'), notebook=False)

# Get the list of personal report files
personal_reports = os.listdir(join(OUTPUT_DIR, "personal_graphs"))
personal_reports, full_names = [f"personal_graphs/{file}" for file in personal_reports], [file.split(".html")[0] for
                                                                                          file in personal_reports]

# Open the personal reports template file and write the links to the personal graphs
with open(join(OUTPUT_DIR, "personal_reports_template.html"), "r", encoding="utf-8") as f:
    file_content = f.read()

# Write the personal reports file
personal_reports_code = ""
for path, name in zip(personal_reports, full_names):
    personal_reports_code += f'\t\t\t<button class="btn btn-info btn-sm mb-2 text-nowrap" onclick="loadHTML(' + f"'{path}'" + f')">{name.replace("_", " ")}</button>\n'

with open(join(OUTPUT_DIR, "personal_reports.html"), "w", encoding="utf-8-sig") as f:
    f.write(file_content.replace("forloop", f"{personal_reports_code}"))


# Add legend to all the graphs

def create_legend(colors, names):
    string = """<div style="position: absolute; top: 20px; right: 20px; z-index: 1000;">
                <div class="card shadow-sm p-2" style="min-width: 180px;">
                <h6 class="text-center mb-2">מקרא</h6>"""
    for i in range(len(colors)):
        string += f"""<div class="d-flex align-items-center mb-1">
                      <span style="background-color: {colors[i]}; width: 16px; height: 16px; display: inline-block; border-radius: 50%; margin-right: 8px;"></span>
                      <span>{names[i]}</span>
                      </div>\n"""
    string += """</div>
              </div>"""
    return string


SCRIPT_FOR_TEXT = """<script>
    // Embedded lines of text instead of reading from .txt file
    const lines = [
      "שורה ראשונה של טקסט",
      "שורה שנייה של טקסט",
    ].join("\\n");

    function updateParagraph() {
      document.getElementById('text').innerText = lines;
    }

    // Start rotation
    updateParagraph();
    setInterval(updateParagraph, 1000); // every 1 seconds
  </script>"""
PARAGRAPH_TEXT = f'<div style="margin-right: 230px;"><p dir="rtl" class="text-end" id="text">טוען...</p></div>'

LEGENDS_DICT = {
    CORRIDOR: ["קומטה שמאל", "מסדרון קומת כניסה", "קומטה ימין"],
    CLASS: ["בופור", "מגן", "סופה", "רעים"],
    SEX: ["גבר", "אישה"],
    SECTOR: ["דתי", "חילוני"],
    BACKGROUND: ["ללא רקע", "פטוריסט", "תאריסט"],
    TIRONUT: ["מחלקה 1", "מחלקה 2"]
}

# Add legend to all the graphs
for population in POPULATIONS:
    filename = FILENAME(population)
    legend = LEGENDS_DICT[population]
    with open(filename, "r") as f:
        content = f.read()

    with open(filename, "w", encoding="UTF-8") as f:
        f.write(content.replace("</body>",
                                f"{SCRIPT_FOR_TEXT}\n{create_legend(COLORS[:len(legend)], legend)}</body>").replace(
            "<body>", f"<body>{PARAGRAPH_TEXT}"))

for filename in os.listdir(join(OUTPUT_DIR, "personal_graphs")):
    with open(join(OUTPUT_DIR, "personal_graphs", filename), "r") as f:
        content = f.read()
    with open(join(OUTPUT_DIR, "personal_graphs", filename), "w", encoding="UTF-8") as f:
        f.write(content.replace("</body>",
                                f'{SCRIPT_FOR_TEXT}\n{create_legend(["red", "green", "purple"], ["קשר שירד", "קשר שהתווסף", "קשר שנשאר"])}</body>').replace(
            "<body>", f"<body>{PARAGRAPH_TEXT}"))
