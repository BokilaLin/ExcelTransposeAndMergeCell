
# coding: utf-8

# In[141]:

import pandas as pd
import numpy as np
import itertools


# In[142]:

df = pd.read_excel('PYTHON.xlsx', sheets_name="工作表1")


# In[143]:

df


# In[144]:

pd.set_option('display.max_columns', 100) # display more columns


# In[145]:

# drop duplicates for each class
df.CLASS1 = df.CLASS1.drop_duplicates() 
df.CLASS2 = df.CLASS2.drop_duplicates()
df.CLASS3 = df.CLASS3.drop_duplicates()


# In[146]:

df.T


# In[147]:

writer = pd.ExcelWriter('new_PYTHON.xlsx', engine='xlsxwriter')
df.T.to_excel(writer, sheet_name='result')


# In[148]:

workbook = writer.book
worksheet = writer.sheets['result'] 
merge_format = workbook.add_format({'align': 'center'}) # align text to center


def findrange(series):
    # return the index of element which is not None
    indexes = []
    for index, value in series.iteritems():
        if pd.notnull(value):
            indexes.append(index)
    indexes.append(len(series)) # append the last index number
    return indexes

def pairwise(iterable):
    """
    return list of tuple of (start, end) index for cell merge from list
    s -> (s0,s1), (s1,s2), (s2, s3), ...
    """
    a, b = itertools.tee(iterable)
    next(b, None)
    return zip(a, b)

# find the merge rule for CLASS1
class1 = df.CLASS1
indexes_1 = findrange(class1)
a = pairwise(indexes_1)
# merge cell on CLASS1
for start, end in a:
    worksheet.merge_range(1, start + 1, 1, end, class1[start], merge_format)


# find the merge rule for CLASS3
class2 = df.CLASS2
indexes_2 = findrange(class2)
b = pairwise(indexes_2)
# merge cell on CLASS3
for start, end in b:
    worksheet.merge_range(2, start + 1, 2, end, class2[start], merge_format)
    


# find the merge rule for CLASS3
class3 = df.CLASS3
indexes_3 = findrange(class3)
c = pairwise(indexes_3)
# merge cell on CLASS3
for start, end in c:
    worksheet.merge_range(3, start + 1, 3, end, class3[start], merge_format)  



# In[150]:

writer.save()


# In[ ]:



