#well, it works, but it doesn't look as good as expected, since the outputs are not
#very precise.

from sklearn.metrics.pairwise import cosine_similarity
import operator
import heapq
import xlwings as xw
# xlwings is used here for it's easy to handle Excel.
import numpy as np

GOOGLE_ENGLISH_WORD_PATH = 'E:/myenglishwords.txt' 
# the file of your phrases
GOOGLE_WORD_FEATURE = 'E:/myenglishwords.vector'
# the vector file you have got from Google's BERT.


def flatten(l, ltypes=(list, tuple)):
    ltype = type(l)
    l = list(l)
    i = 0
    while i < len(l):
        while isinstance(l[i], ltypes):
            if not l[i]:
                l.pop(i)
                i -= 1
                break
            else:
                l[i:i + 1] = l[i]
        i += 1
    return ltype(l)

try:
    import cPickle
except :
    import _pickle as cPickle


def save_model(clf,modelpath): 
    with open(modelpath, 'wb') as f: 
        cPickle.dump(clf, f) 
        
def load_model(modelpath): 
    try: 
        with open(modelpath, 'rb') as f: 
            rf = cPickle.load(f) 
            return rf 
    except Exception as e:        
        return None 

phrase_model=load_model(GOOGLE_WORD_FEATURE)
print(len(phrase_model))
print(phrase_model.keys())


def phrase_similarity(phrase, N=10):
    phrase_vec = phrase_model[phrase]
    CosDisList = []
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open('someexample.xls')
    sht = wb.sheets['sheet1']
    
    
    for a_word in phrase_model.keys():

        a_val = phrase_model[a_word]
        cos_dis = cosine_similarity(phrase_vec, a_val)

        for i in range(1, 3734):

            if a_word == sht.cells(i, 1).value:
                DataFromExcel = (sht.cells(i, 2).value, sht.cells(i, 3).value, sht.cells(i, 4).value, sht.cells(i, 5).value, sht.cells(i, 6).value)
                DataCombined = (a_word, DataFromExcel)
                CosDisBind = [float(str(cos_dis.tolist()).strip('[[]]')), DataCombined]

                CosDisList.append(CosDisBind)

                CosDisListSort = sorted(CosDisList, key=operator.itemgetter(0), reverse=True)

                CosDisListTopN = heapq.nlargest(N, CosDisListSort)
                
                flat_list = flatten(CosDisListTopN)
                
                final_array = np.array(flat_list).reshape((np.round(len(flat_list)//7), 7)).tolist()
                
                most_final = ["\t".join(x) for x in final_array]
            

    for EveryItem in most_final:
        print(EveryItem)


                

    
phrase_similarity('iphone', 10)  
#search phrases that are similar to the phrase "iphone", and only output 10 ones.