import pandas as pd
import collections
import re
import csv
#reading excel data file
#and extracting the id coulmn only
xl = pd.read_excel('ZCEC14.xls', usecols=[0])
for name, val in xl.iteritems():
    pass

#storing the id in a list for further analysis
orf_id_xl = []
for i in range(len(val)):
    orf_id_xl.append(val[i])

#reading fasta file and storing it into a dict
#of id_seq pairs
with open('4-All Orfs protein seq.fa', 'r+') as handle:
    x = 0
    pair_id_seq = collections.defaultdict(str)
    while True:
        try:
            seq = handle.readline()
            if seq[0] == ">":
                hold = seq
                pair_id_seq[seq] = handle.readline()
            else:
                pair_id_seq[hold] = pair_id_seq[hold] + seq
        except:
            break
#initializing the matched ids with its corrospending sequences
new_seq_excel_row = []
matched_orfs = []

for key, value in pair_id_seq.items():
    terminate_index = key[5:11].find(":")
    if terminate_index == -1:
        new_key = key[5:11]
    elif terminate_index != -1:
        new_key = key[5:terminate_index + 5]
    for i in orf_id_xl:
        if new_key == i:
            print(new_key, i)
            matched_orfs.append(i)
            new_seq_excel_row.append(value)
            break
#final dict of pairs of matched ids and seq
final_result = dict(zip(matched_orfs, new_seq_excel_row))
print(len(final_result))

#exporting the dictionary into excel file
with open('final.csv', 'w') as output:
    excel_file = csv.writer(output)
    for key, value in final_result.items():
        excel_file.writerow([key, value])