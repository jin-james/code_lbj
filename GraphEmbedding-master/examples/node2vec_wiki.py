import numpy as np
from function import field_to_json
from ge.classify import read_node_label
from ge import Node2Vec
from time import time
import matplotlib.pyplot as plt
import networkx as nx
from sklearn.manifold import SpectralEmbedding, Isomap, TSNE, MDS, LocallyLinearEmbedding
from sklearn.decomposition import PCA, KernelPCA, LatentDirichletAllocation as LDA

# def evaluate_embeddings(embeddings):
#     X, Y = read_node_label('../data/wiki/wiki_labels.txt')
#     tr_frac = 0.8
#     print("Training classifier using {:.2f}% nodes...".format(
#         tr_frac * 100))
#     clf = Classifier(embeddings=embeddings, clf=LogisticRegression())
#     clf.split_train_evaluate(X, Y, tr_frac)
#
#
def plot_embeddings(embeddings,):
    X, Y = read_node_label('../data/wiki/karate.txt')
    emb_list = []
    embeddings = field_to_json(embeddings)
    key = embeddings.keys()

    for k in key:
        emb_list.append(embeddings[k])
    emb_list = np.array(emb_list)

    # model = TSNE(n_components=2)
    # node_pos = model.fit_transform(emb_list)

    # color_idx = {}
    # for i in range(len(X)):
    #     color_idx.setdefault(Y[i][0], [])
    #     color_idx[Y[i][0]].append(i)
    #
    # for c, idx in color_idx.items():
    #     plt.scatter(node_pos[idx, 0], node_pos[idx, 1], label=c)
    for i in range(len(key)):
        plt.scatter(emb_list[i, 0], emb_list[i, 1], alpha=0.5, s=150, color='b')
        plt.text(emb_list[i, 0], emb_list[i, 1], list(key)[i], horizontalalignment='center', verticalalignment='center')
        # for j in range(len(X)):
        #     if X[j] == str(i+1):
        #         x = int(X[j])
        #         y = int(Y[j][0])
        #         plt.plot(node_pos[int(list(key)[x-1])-1], node_pos[int(list(key)[y-1])-1], color='r')
    plt.legend()
    plt.show()


if __name__ == "__main__":
    t0 = time()
    # parser.add_argument('--output', required=True,
    #                     help='Output representation file')
    G = nx.read_edgelist('../data/wiki/karate.txt',
                         create_using=nx.DiGraph(), delimiter='\t', nodetype=None, data=[('weight', int)])

    model = Node2Vec(G, walk_length=10, num_walks=40,
                   p=0.25, q=4, workers=1)
    model.train(window_size=5, iter=3)
    embeddings = model.get_embeddings()
    # evaluate_embeddings(embeddings)
    with open('../output', 'w') as f:
        print(embeddings, f)
    plot_embeddings(embeddings)
    t1 = time()
    print('it works {}s'.format(t1-t0))
