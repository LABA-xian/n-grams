# -*- coding: utf-8 -*-
"""
Created on Sat Apr 28 17:19:25 2018

@author: MB207-1
"""

from nltk import FreqDist
from nltk import ngrams
import pandas as pd
import collections
import csv
import re

word = []

df = pd.read_excel('C:/Users/MB207-1/Desktop/統一速達_專業QA_171213.xlsx')

Qlist = (df['延伸問法'].astype(str)).tolist()

for q in Qlist:
    if q == ' ' or q == 'nan':
        print('skip')
    else:
        q = re.sub('\W', '', q)
        bigrams = ngrams(q, 4)
        
        bigramsDist = FreqDist(bigrams)
        aa = list(bigramsDist)
        n_list =  bigramsDist.most_common(10)
        for n in n_list:
            #print(''.join(list(n[0])))
            word.append(''.join(list(n[0])))

frequence = list(collections.Counter(word).items())


with open('output4.csv', 'w', encoding="utf_8_sig", newline='') as csvfile:
    writer = csv.writer(csvfile)
    for f in frequence:
        writer.writerow([f[0], f[1]])

print("\各詞出現的次數為：\n %s" % collections.Counter(word))
    
    
