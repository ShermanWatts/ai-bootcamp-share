"""
Example from discussion from class - use dict to lookup encoded values (to decode)
"""

import pandas as pd

#csv-like data
raw_data = [
    ['gender','education'],
    [0,0],
    [1,1],
    [0,2],
    [1,2]
]
column_headers = raw_data[0] #first row
rows_data = raw_data[1:] #rest of rows

#create a DataFrame
df = pd.DataFrame(raw_data,columns=column_headers)

#or load the DataFrame from a csv
#df = pd.read_csv("data.csv")

#lookup dict to decode values
lookup = {
    'gender': {0: 'Male', 1: 'Female'},
    'education': {0: 'High School', 1: 'College', 2: 'Post Graduate'}
}

#do some processing and lookup actual values if you ever need to
for row_num in range(1, len(df)):  # Start from index 1 instead of 0
    gender_val = df.iloc[row_num, 0] 
    education_val = df.iloc[row_num, 1] 

    # Encoded values
    print(f"Row {row_num}: {gender_val}, {education_val}")

    # Decode values using the lookup dictionary
    gender_text = lookup[df.columns[0]][gender_val]  # Lookup for 'gender'
    education_text = lookup[df.columns[1]][education_val]  # Lookup for 'education'
    
    print(f"Row {row_num}: {gender_text}, {education_text}")