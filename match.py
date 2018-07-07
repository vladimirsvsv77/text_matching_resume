from vacancies import vacancy_dict

from gensim.models.wrappers import FastText
import gensim
import pandas as pd
from scipy.spatial.distance import euclidean

from openpyxl import load_workbook
import string

model = FastText.load_fasttext_format('data.bin')
wb = load_workbook('candidates.xlsx')


def clean_str(s):
    for c in string.punctuation:
        s = s.replace(c, "")
    return s


def get_similarity_euql(model, first_sentence, second_sentence):
    similarity = 0
    first_sentence = [i for i in clean_str(first_sentence).split() if i in model]
    second_sentence = [i for i in clean_str(second_sentence).split() if i in model]
    for i in first_sentence:
        first_vector = model[i]
        sim_i = 0
        for j in second_sentence:
            second_vector = model[j]
            sim_i += euclidean(first_vector, second_vector)
        similarity += sim_i / len(second_sentence)
    return similarity / len(first_sentence)


class Candidate():
    
    def __init__(self, name, phone, email, text, vector_1, vector_2, 
                 vector_3, vector_4, vector_5, vector_6, vector_7, vector_8, vector_9, vector_10):
        self.name = name
        self.phone = phone
        self.email = email
        self.text = text
        self.vector_1 = vector_1
        self.vector_2 = vector_2
        self.vector_3 = vector_3
        self.vector_4 = vector_4
        self.vector_5 = vector_5
        self.vector_6 = vector_6
        self.vector_7 = vector_7
        self.vector_8 = vector_8
        self.vector_9 = vector_9
        self.vector_10 = vector_10
    
    def to_dict(self):
        return {
            'name': self.name,
            'phone': self.phone,
            'email': self.email,
            'text': self.text,
            'vector_1': self.vector_1,
            'vector_2': self.vector_2,
            'vector_3': self.vector_3,
            'vector_4': self.vector_4,
            'vector_5': self.vector_5,
            'vector_6': self.vector_6,
            'vector_7': self.vector_7,
            'vector_8': self.vector_8,
            'vector_9': self.vector_9,
            'vector_10': self.vector_10,
        }


candidates = []

for i in range(rows):
    if i >= 2:
        cand = Candidate(sheet_ranges['B' + str(i)].value, 
                         sheet_ranges['U' + str(i)].value, 
                         sheet_ranges['T' + str(i)].value, 
                         str(str(sheet_ranges['G' + str(i)].value) + '. ' +
                         str(sheet_ranges['J' + str(i)].value) + '. ' +
                         str(sheet_ranges['K' + str(i)].value) + '. ' + 
                         str(sheet_ranges['L' + str(i)].value) + '. ' + 
                         str(sheet_ranges['M' + str(i)].value) + '. ' + 
                         str(sheet_ranges['N' + str(i)].value) + '. ' + 
                         str(sheet_ranges['O' + str(i)].value) + '. ' + 
                         str(sheet_ranges['P' + str(i)].value) + '. ' + 
                         str(sheet_ranges['Q' + str(i)].value) + '. ' + 
                         str(sheet_ranges['R' + str(i)].value) + '. ' + 
                         str(sheet_ranges['S' + str(i)].value)).replace('None', ''),
                        0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0)
        candidates.append(cand) 


df = pd.DataFrame.from_records([c.to_dict() for c in candidates])


for vac_number in range(11):
    if vac_number >= 1:
        vacancy_text = vacancy_dict['vacancy_' + str(vac_number)].replace('\n', '').replace('•', '')
        for i, row in df.iterrows():
            try:
                df.at[i,'vector_' + str(vac_number)] = get_similarity_euql(model, vacancy_text, df.at[i,'text'])
            except:
                df.at[i,'vector_' + str(vac_number)] = 100