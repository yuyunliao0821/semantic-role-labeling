# -*- coding: utf-8 -*-
import spacy
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import sys
from collections import OrderedDict

non_bmp_map = dict.fromkeys(range(0x10000, sys.maxunicode + 1), 0xfffd)

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

# Find a workbook by name and open the first sheet
# Make sure you use the right name here.
sh = client.open("Syntax - Interns")
worksheet = sh.worksheet('sentences')
worksheet2 = sh.worksheet('Syntax')

# Extract sentences
sentlist = worksheet.col_values(4)[1:]

# Extract matches
match = worksheet2.col_values(2)[2:]
symbols = worksheet2.col_values(3)[2:]
nlp = spacy.load('en_core_web_sm')


# remove duplicates in a string
def removeDupWithOrder(str):
    return "".join(OrderedDict.fromkeys(str))


for record in sentlist:
    sent = record

    doc = nlp(sent)

    reducedlist = []
    lemlist = [i.lemma_ for i in doc]
    textlist = [i.text for i in doc]
    poslist = [i.pos_ for i in doc]
    taglist = [i.tag_ for i in doc]
    deplist = [i.dep_ for i in doc]

    for i in range(len(doc)):

        if 'not' in lemlist:
            reducedlist.append('not')

        if poslist[i] == 'NOUN':
            reducedlist.append(lemlist[i])

        if taglist[i] == 'WRB' and lemlist[i + 1] == 'do':  # how do
            a = lemlist[i], lemlist[i + 1]
            verbcomp = ' '.join(a)
            reducedlist.append(verbcomp)
        elif taglist[i] == 'WRB':  # how
            reducedlist.append(lemlist[i])

        if taglist[i] == 'RB':
            reducedlist.append(lemlist[i])

        if poslist[i] == 'CCONJ':
            reducedlist.append(lemlist[i])

        if poslist[i] == 'VERB' and taglist[i + 1] == 'IN':  # go to
            a = lemlist[i], lemlist[i + 1]
            verbcomp = ' '.join(a)
            reducedlist.append(verbcomp)
        elif poslist[i] == 'VERB' and poslist[i + 1] == 'PUNCT':
            reducedlist.append(lemlist[i])
        elif poslist[i] == 'VERB' and taglist[i + 2] == 'IN':  # help A with
            a = lemlist[i], lemlist[i + 2]
            verbcomp = ' '.join(a)
            reducedlist.append(verbcomp)
            reducedlist.append(lemlist[i])
        elif poslist[i] == 'VERB' and taglist[i + 1] == 'TO':  # like to, need to
            a = lemlist[i], lemlist[i + 1]
            verbcomp = ' '.join(a)
            reducedlist.append(verbcomp)
        elif poslist[i] == 'VERB' and taglist[i + 1] == 'NN':  # have breakfast
            b = lemlist[i], lemlist[i + 1]
            verbcomp = ' '.join(b)
            reducedlist.append(verbcomp)

        if poslist[i] == 'VERB' and deplist[i] == 'aux':  # have, do
            pass
        elif poslist[i] == 'VERB' and taglist[i] == 'MD':
            reducedlist.append(lemlist[i])
        elif poslist[i] == 'VERB':  # play soccer, like music
            reducedlist.append(lemlist[i])

        if taglist[i] == 'JJ' and taglist[i + 1] == 'IN':  # interested in
            b = lemlist[i], lemlist[i + 1]
            verbcomp = ' '.join(b)
            reducedlist.append(verbcomp)
        elif taglist[i] == 'JJ':  # old, young, favorite
            reducedlist.append(lemlist[i])

        if taglist[i] == 'IN' and taglist[i + 1] == 'NN':  # at home, at school
            b = lemlist[i], lemlist[i + 1]
            verbcomp = ' '.join(b)
            reducedlist.append(verbcomp)

        if taglist[i] == 'IN' and poslist[i + 1] == 'PROPN':
            b = taglist[i], poslist[i + 1]
            verbcomp = ''.join(b)
            reducedlist.append(verbcomp)

        if lemlist[i] == 'there' and lemlist[i+1] == 'be':
            b = lemlist[i], lemlist[i + 1]
            verbcomp = ' '.join(b)
            reducedlist.append(verbcomp)

        if lemlist[i] == 'if' or lemlist[i] == 'If':
            reducedlist.append(lemlist[i])

        if lemlist[i] == 'whether' or lemlist[i] == 'Whether':
            reducedlist.append(lemlist[i])

        if lemlist[i] == 'Where':
            reducedlist.append(lemlist[i])

        if textlist[i] == '—':
            reducedlist.append(textlist[i])

    verblist = list(dict.fromkeys(reducedlist))  # remove duplicates

    symlist = []

    for i in range(len(verblist)):

        if verblist[i] in match:
            symlist.append(symbols[match.index(verblist[i])])
            outputstr = ''.join(symlist)
            outputstr = outputstr.replace('(', '')
            outputstr = outputstr.repl8ace(')', '')

    if len(symlist) == 0:
        print("NO MATCH")

    else:
        print(removeDupWithOrder(outputstr))
