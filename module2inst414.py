import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt


got = pd.read_excel("01 A Game Of Thrones.xlsx")
cok = pd.read_excel("02 A Clash Of Kings.xlsx")
sos = pd.read_excel("03 A Storm Of Swords.xlsx")
ffc = pd.read_excel("04 A Feast  For Crows.xlsm")
dwd = pd.read_excel("05 A Dance With Dragons.xlsx")
books = [got, cok, sos, ffc, dwd]
new_books = []
row_drops = ["CHAPTER", "EPILOGUE", "PROLOGUE"]

chapter_mentions = {}
#Drops unneccesary rows and columns from the dataframes
for book in books[:3]:
    new_book = book.drop(columns=['Debut', 'Page In', 'Page Out'])    
    new_book = new_book[~new_book['Character'].str.contains('|'.join(row_drops), case=False, na=False)]
    new_book = new_book[~new_book['Character'].str.contains('|'.join(row_drops))]
    new_books.append(new_book)
    

for book in books[3:]:
    new_book = book.drop(columns=['Debut', 'Demise'])
    new_book = new_book[~new_book['Character'].str.contains('|'.join(row_drops))]
    new_books.append(new_book)
   

g = nx.Graph()

#Add nodes
for book in new_books:
    for index, row in book.iterrows():
        if book['Character'][index] not in g.nodes:
            g.add_node(book['Character'][index].strip())
        chapter_mentions[book['Character'][index].strip()] = chapter_mentions.get(book['Character'][index].strip(), 0) + 1

            
#Add edges
for book in new_books:
    column_names = book.columns.tolist()
    for index, row in book.iterrows():
        for column in column_names:
            if column != 'Character':
                if (type(book[column][index]) == str and
                chapter_mentions.get(book['Character'][index].strip(), 0) >= 5 and
                chapter_mentions.get(book[column][index].strip(), 0) >= 5):
                    if not g.has_edge(book['Character'][index].strip(), book[column][index].strip()):
                        g.add_edge(book['Character'][index].strip(), book[column][index].strip(), weight=1)
                    else:
                        g[book['Character'][index].strip()][book[column][index].strip()]['weight'] += 1
g.remove_nodes_from(list(nx.isolates(g)))

g = g.subgraph(max(nx.connected_components(g), key=len)).copy()


nx.write_gexf(g, "asoiaf_network.gexf")

betweenness = nx.betweenness_centrality(g, normalized=True)
PageRank = nx.pagerank(g)
degree = nx.degree_centrality(g)


centrality_df = pd.DataFrame({
    'Character': list(betweenness.keys()),
    'Betweenness Centrality': list(betweenness.values()),
    'PageRank': list(PageRank.values()),
    'Degree': list(degree.values())
})

important_betweeness = centrality_df.sort_values(by='Betweenness Centrality', ascending=False)
important_pagerank = centrality_df.sort_values(by='PageRank', ascending=False)

print("Betweenness Centrality Head")
print(important_betweeness.head(10))
print("PageRank Head")
print(important_pagerank.head(10))

# Export to CSV
centrality_df.to_csv('character_centrality.csv', index=False)

from networkx.algorithms.components import connected_components

# Step 1: Get the top 100 characters by PageRank
top_100_characters = important_pagerank['Character'].head(100).tolist()

# Step 2: Create a subgraph with only the top 100 characters
top_100_subgraph = g.subgraph(top_100_characters).copy()

# Step 3: Keep only the largest connected component
largest_cc = max(connected_components(top_100_subgraph), key=len)
main_subgraph = top_100_subgraph.subgraph(largest_cc).copy()

# Function to draw the graph using different layouts
def draw_graph(graph, layout_func, layout_name, filename):
    plt.figure(figsize=(15, 15))
    pos = layout_func(graph, k=.9)

    # Draw nodes
    node_sizes = [PageRank[node] * 10000 for node in graph.nodes()]
    nx.draw_networkx_nodes(graph, pos, node_size=node_sizes, node_color='skyblue', alpha=0.8)

    # Draw edges
    nx.draw_networkx_edges(graph, pos, edge_color='gray', alpha=0.3)

    # Draw labels
    nx.draw_networkx_labels(graph, pos, font_size=14, font_color='#333333')

    plt.title("Top 100 Characters by PageRank", fontsize=18)
    plt.axis('off')
    plt.tight_layout()
    plt.savefig(filename, format='png', dpi=300)

    plt.show()

# Try different layouts
draw_graph(main_subgraph, nx.spring_layout, "Spring Layout", "top100_spring_layout.png")

