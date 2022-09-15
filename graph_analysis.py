import scipy as sp
import networkx as nx
import matplotlib.pyplot as plt
import xlrd
from networkx.algorithms import community
file = "./tweets2.xlsx"
G = nx.Graph()
names = []

name = xlrd.open_workbook(file)
sheet = name.sheet_by_index(0)

for row in range(1,sheet.nrows):
    data = sheet.row_slice(row)
    person1 = data[0].value
    person2 = data[1].value
    names.append((person1 , person2))

G.add_edges_from(names)
print("Degree Centrality",nx.degree_centrality(G))
print("Closensess Centrality",nx.closeness_centrality(G, u=None, distance=None))
print("Harmonic Centrality",nx.harmonic_centrality(G, nbunch=None, distance=None, sources=None))
print("The Maximal Matching",sorted(nx.maximal_matching(G)))

communities_generator = community.girvan_newman(G) #Finds communities in a graph using the Girvan–Newman method.
top_level_communities = next(communities_generator)
next_level_communities = next(communities_generator)
nx.draw(G)
plt.show()