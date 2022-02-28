import base64
from github import Github
from github import InputGitTreeElement
import pyarrow.parquet as pq
import numpy as np
import pandas as pd
import pyarrow as pa
import re
import requests
df = pd.read_csv(r"C:/Users/spasikhani/Documents/mass production test/2015 NGBS Project data extraction.csv")
while True:
    try:
        pqtable = pa.Table.from_pandas(df)
        break
    except TypeError as err:
        pattern = "column (.*?) with"
        print(type(err))
        column = re.search(pattern, str(err)).group(1)
        df[column] = df[column].astype('string')

pq.write_table(pqtable, 'C:/Users/spasikhani/Documents/mass production test/NGBS2015data.parquet', compression='snappy')



repo_dir =
user = 'bsohrab'
password = "S!s1s1s1"
token = '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'
g = requests.get('https://api.github.com/search/repositories?q=github+api', auth=(user,token))
g = Github(token)

repo = g.get_user().get_repo("bsohrab/NGBS_Receptacle")

file_list = ['C:/Users/spasikhani/Documents/mass production test/NGBS2015data.parquet']
file_names = ['NGBS2015data.parquet']

commit_message = 'the light warrior'
master_ref = repo.get_gif_ref('heads/master')
master_sha = master_ref.object.sha
base_tree  = repo.get_git_tree(master_sha)

element_list = list()

for i, entry in enumerate(file_list):
    with open(entry) as input_file:
        data = input_file.read()
    if entry.endswith(".png"):
        data = base64.encode((data))
    element = InputGitTreeElement(file_names[i], '100644', 'blob', data)
    element_list.append((element))
tree = repo.create_git_tree(element_list, base_tree)
parent = repo.get_git_commit(master_sha)
commit = repo.create_git_commit(commit_message, tree, [parent])
master_ref.edit(commit.sha)
