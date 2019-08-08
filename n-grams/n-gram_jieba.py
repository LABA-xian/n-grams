# -*- coding: utf-8 -*-
"""
Created on Tue May  8 20:03:46 2018

@author: MB207-1
"""

from nltk import FreqDist
from nltk import ngrams
import pandas as pd
import collections
import jieba.posseg as pseg
import re
import csv

word = []
wordlist = []


n_gram_dict = {}

df = pd.read_excel('C:/Users/MB207-1/Desktop/統一速達_專業QA_171213.xlsx')

Qlist = (df['延伸問法'].astype(str)).tolist()

for i in range(5):
    word = []
    for q in Qlist:
        if q == ' ' or q == 'nan':
            print('skip')
        else:
            q = re.sub('\W', '', q)
            bigrams = ngrams(q, i + 2)
            bigramsDist = FreqDist(bigrams)
            n_list =  bigramsDist.most_common(10)
            for n in n_list:
                word.append(''.join(list(n[0])))
    n_gram_dict[i + 2] = word

j = 0

for key, values in n_gram_dict.items():
    frequence = collections.Counter(values)
    top50 = list(frequence.most_common(200))
    with open(str(j) + '.csv', 'w', encoding="utf_8_sig", newline='') as csvfile:
        writer = csv.writer(csvfile)
        for top,vlues in top50:
            wordcut = pseg.cut(top)
            for wordpseg, flag in wordcut:
                writer.writerow([wordpseg, flag])
    j = j + 1