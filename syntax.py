import spacy
import xlsxwriter
import xlrd

data=[]
book = xlrd.open_workbook('hello.xlsx')
sheet = book.sheet_by_name('sheet2')

for i in range(504):
    data+=sheet.row_values(i)

workbook = xlsxwriter.Workbook('Syntaxoutput.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0

nlp = spacy.load('en_core_web_sm')

for record in data:
 sent = record

 doc = nlp(sent)

 deplist = []
 sentlist = []
 textlist = []
 for i in doc:
    deplist.append(i.dep_)
    sentlist.append(i.lemma_)
    textlist.append(i.text)

 Arg=[]

# 計算受詞量
 countobj = 0
 for i in deplist:
     if i == 'pobj' or i == 'dobj':
        countobj += 1


# 計算prep phrase量
 countpp = 0
 pp = []
 if len(deplist)>6:
     
  for i in range(len(deplist)):

     if deplist[i] == 'prep' and deplist[i+1] == 'punct':
         pass

     elif deplist[i] == 'prep' and deplist[i + 2] == 'pobj':
        pptuple = (textlist[i], textlist[i + 2]) #to the manager
        pp.append(pptuple)
        
         


 if countobj >= 2:

    print (data.index(record))

    for e in range(len(sentlist)):

        if deplist[e] == 'nsubj':
            if deplist[e-1]== 'compound':
                substr = textlist[e-1], textlist[e]
                sub = ' '.join(substr)
                Arg.append(sub)
        
            elif textlist[e]=='that':
                pass
            else:
                sub = textlist[e]
                Arg.append(sub)
                
        elif deplist[e] == 'attr' and deplist[e-3] == 'expl':
            sub = textlist[e]
            Arg.append(sub)
    
        

        if (deplist[e] == 'ROOT' or deplist[e]=='advcl' or deplist[e]=='ccomp'
            or deplist[e]=='conj' or deplist[e]=='relcl'):

            if 2 * len(pp) > 2:

                if deplist[e+1] == 'punct': #...were damaged
                    verb= textlist[e]
                    Arg.append(verb)
                    
                elif deplist[e+2] == 'xcomp': #want to...
                    for record in pp:
                        verbstr= textlist[e+2], record[0]
                        verb= ' '.join(verbstr)
                        Arg.append(verb)
                else:
                    for record in pp:
                        verbstr = textlist[e], record[0]
                        verb = ' '.join(verbstr)
                        Arg.append(verb)
                

            else:

                if deplist[e + 1] == 'prt' or deplist[e + 1] == 'prep':
                    ee = textlist[e], textlist[e + 1]
                    verb = ' '.join(ee)
                    Arg.append(verb)

                elif deplist[e+1]=='punct' or deplist[e+2]=='punct':
                    verb = textlist[e]
                    Arg.append(verb)

                elif deplist[e + 2] == 'prt' or deplist[e + 2] == 'prep':
                    a = textlist[e], textlist[e + 2]
                    verb = ' '.join(a)
                    Arg.append(verb)
                elif deplist[e+2]=='xcomp' and deplist[e+3]=='punct': #Do you like to read?
                    verb = textlist[e+2]
                    Arg.append(verb)
                elif deplist[e + 2] == 'xcomp' and deplist[e + 4] == 'prep': #want to help Mary with
                    c = textlist[e + 2], textlist[e + 4]
                    verb = ' '.join(c)
                    Arg.append(verb)
                elif deplist[e + 2] == 'xcomp' and deplist[e + 4] == 'prt':
                    f = textlist[e + 2], textlist[e + 4]
                    verb = ' '.join(f)
                    Arg.append(verb)
                elif deplist[e + 2] =='xcomp' and deplist[e+4] == 'dobj': #want to reserve a table
                    verb = textlist[e]
                    Arg.append(verb)
                elif deplist[e + 2] == 'xcomp' and deplist[e + 3] == 'prep': # want to go to
                    g = textlist[e + 2], textlist[e + 3]
                    verb = ' '.join(g)
                    Arg.append(verb)
                elif deplist[e + 2] == 'xcomp' and deplist[e + 3] == 'prt':
                    h = textlist[e + 2], textlist[e + 3]
                    verb = ' '.join(h)
                    Arg.append(verb)
                elif deplist[e + 2] == 'xcomp' and deplist[e + 3] == 'dative':  # John wants to teach me math.
                    verb = textlist[e + 2]
                    Arg.append(verb)

                elif deplist[e + 3] == 'prep':
                    j = textlist[e], textlist[e + 3]
                    verb = ' '.join(j)
                    Arg.append(verb)
                else:
                    verb = textlist[e]
                    Arg.append(verb)


        if deplist[e] == 'dobj':
            
            if deplist[e - 2] == 'prt' or deplist[e - 2] == 'prep':
                obj1 = textlist[e]
                Arg.append(obj1)

            elif deplist[e+1] == 'punct' or deplist[e+2]=='punct':
                obj = textlist[e]
                Arg.append(obj)

            elif deplist[e - 1] == 'dative':
                d = textlist[e - 1], textlist[e]
                obj1 = ' - '.join(d)
                Arg.append(obj1)

            elif deplist[e + 1] == 'prt':
                obj1 = textlist[e]
                Arg.append(obj1)

            elif deplist[e + 2] == 'pobj':

                obj1 = textlist[e]
                Arg.append(obj1)

            elif deplist[e + 3] == 'pobj':
                obj1 = textlist[e]
                Arg.append(obj1)

            elif deplist[e-1]== 'advcl':
                pass

            else:
                obj1 = textlist[e]
                Arg.append(obj1)

       
        if deplist[e] == 'pobj':
            if deplist[e + 1] == 'mark':
                obj2 = textlist[e]
                Arg.append(obj2)
            else:
                obj2 = textlist[e]
                Arg.append(obj2)
                
                
                
                
    b= str(Arg)
    worksheet.write(row, col, b)
    row+=1
    


 elif countobj <= 1:

    print (data.index(record)) 

    for e in range(len(sentlist)):


        if e>0 and deplist[e] == 'nsubj':
            if deplist[e-1]== 'compound':
                substr = textlist[e-1], textlist[e]
                sub = ' '.join(substr)
                Arg.append(sub)
        
        
            else:
                sub = textlist[e]
                Arg.append(sub)

        if deplist[e]== 'nsubjpass':
            sub= textlist[e]
            Arg.append(sub)
                

        if (deplist[e] == 'ROOT' or deplist[e]=='advcl' or deplist[e]=='ccomp'
            or deplist[e]=='conj' or deplist[e]=='relcl'):


            if deplist[e-2]=='ROOT' or deplist[e-3]=='ROOT':
                pass
    
            elif deplist[e+1]== 'punct':
                verb = textlist[e]
                Arg.append(verb)
                
            elif deplist[e+1]=='punct' or deplist[e+2]=='punct':
                verb = textlist[e]
                Arg.append(verb)

            elif deplist[e+2]=='xcomp' and deplist[e+3]=='punct': #Do you like to read?
                verb = textlist[e+2]
                Arg.append(verb)

            elif deplist[e + 2] == 'xcomp' and deplist[e + 4] == 'prep':
                c = textlist[e + 2], textlist[e + 4]
                verb = ' '.join(c)
                Arg.append(verb)

            elif deplist[e + 2] == 'xcomp' and deplist[e + 4] == 'prt':
                f = textlist[e + 2], textlist[e + 4]
                verb = ' '.join(f)
                Arg.append(verb)

            elif deplist[e + 2] == 'xcomp' and deplist[e + 3] == 'prep':
                g = textlist[e + 2], textlist[e + 3]
                verb = ' '.join(g)
                Arg.append(verb)
            elif deplist[e + 2] == 'xcomp' and deplist[e + 3] == 'prt':
                h = textlist[e + 2], textlist[e + 3]
                verb = ' '.join(h)
                Arg.append(verb)
            elif deplist[e + 2] == 'xcomp' and deplist[e + 3] == 'dative':  # John wants to teach me math.
                verb = textlist[e + 2]
                Arg.append(verb)
            elif deplist[e+2] == 'xcomp' and deplist[e+1] == 'aux':
                verb = textlist[e+2]
                Arg.append(verb)
            elif deplist[e + 1] == 'prt' or deplist[e + 1] == 'prep':
                ee = textlist[e], textlist[e + 1]
                verb = ' '.join(ee)
                Arg.append(verb)

            elif deplist[e + 2] == 'prt' or deplist[e + 2] == 'prep':
                a = textlist[e], textlist[e + 2]
                verb = ' '.join(a)
                Arg.append(verb)

            elif deplist[e + 3] == 'prt' or deplist[e + 3] == 'prep':
                a = textlist[e], textlist[e + 2]
                verb = ' '.join(a)
                Arg.append(verb)
            elif deplist[e + 1] == 'dobj':
                verb = textlist[e]
                Arg.append(verb)
            else:
                verb = textlist[e]
                Arg.append(verb)

        if deplist[e] == 'pobj':
            if deplist[e - 1] == 'compound':
                ab = textlist[e - 1], textlist[e]
                obj = ' '.join(ab)
                Arg.append(obj)
            else:
                obj = textlist[e]
                Arg.append(obj)

        if deplist[e] == 'dative':
                obj = textlist[e]
                Arg.append(obj)

        if deplist[e] == 'dobj':
                obj = textlist[e]
                Arg.append(obj)
                
        if deplist[e] == 'attr' or deplist[e] == 'acomp':
            obj = textlist[e]
            Arg.append(obj)

        if deplist[e]== 'npadvmod':
            time = textlist[e]
            Arg.append(time)
            
    a= str(Arg)
    worksheet.write(row, col, a)
    row+=1

workbook.close()

# Put the book on the table onto the shelf.  - put onto - book on the table - shelf
# I will take tomorrow off if he is off.
# I'd like to complain to the manager about your service. I - complain to - manager. I - complain about - service
# The rock star needs a bodyguard to protect him from the crazy fans.
# multiple sentences with prep phrases.
# observations, evidentiality(may have happened, could,...seem to be, appear to be)

