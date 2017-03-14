#!/usr/bin/python
# -*- coding: utf-8 -*-

import csv
import getopt
import json
import sys
import unicodecsv
import xlrd
import json
import pandas as pd
import xlrd
from itertools import groupby
from collections import OrderedDict


# Store input and output file names
ifile=''
ofile=''


try:
    myopts, args = getopt.getopt(sys.argv[1:],"i:")
except getopt.GetoptError as e:
    print (str(e))
    print("Usage: %s -i input" % sys.argv[0])
    sys.exit(2)


###############################
# o == option
# a == argument passed to the o
###############################
for o, a in myopts:
    if o == '-i':
        ifile=a
    else:
        print("Usage: %s -i input" % sys.argv[0])

if not ifile:
    print("Usage: %s -i input" % sys.argv[0])
    sys.exit(2)

## transform xslx into csv
book = xlrd.open_workbook(ifile)
sh = book.sheet_by_index(0)
tocsv = open(ifile+'.csv', 'w')
wr = unicodecsv.writer(tocsv, encoding='utf-8', quoting=unicodecsv.QUOTE_ALL)

for rownum in xrange(sh.nrows):
        wr.writerow(sh.row_values(rownum))

tocsv.close()

## define a new result from csv parsing
df = pd.read_csv(ifile +'.csv', dtype={
            "Région" : str,
            "ancienne région" : str,
            "N° Département" : str,
            "Département" : str,
            "Nom" : str,
            "Prénom" : str,
            "E-Mail" : str
        })

results = []

for (region), bag in df.groupby(["Région"]):
    contents_df = bag.drop(["Région"], axis=1)
    subset = [OrderedDict(row) for i,row in contents_df.iterrows()]
    results.append(OrderedDict([("Région", region),
                                ("Contacts", subset)]))

#print json.dumps(results[0], indent=4)
with open(ifile+'.json', 'w') as outfile:
    outfile.write(json.dumps(results, indent=4))