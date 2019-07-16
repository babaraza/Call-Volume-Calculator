from pathlib import Path
import pandas as pd

data = [{
    2018: {
    'Jan':{'incoming': 10, 'outgoing': 20},
    'Feb':{'incoming': 30, 'outgoing': 40}},
    2019: {
    'Jan':{'incoming': 50, 'outgoing': 60},
    'Feb':{'incoming': 70, 'outgoing': 80}}
}]

data2 = [{'posts': {'item_1': 1,
                   'item_2': 8,
                   'item_3': 105,
                   'item_4': 324,
                   'item_5': 313, }
         },
        {'edits': {'item_1': 1,
                   'item_2': 8,
                   'item_3': 61,
                   'item_4': 178,
                   'item_5': 163}
         },
        {'views': {'item_1': 2345,
                   'item_2': 330649,
                   'item_3': 12920402,
                   'item_4': 46199102,
                   'item_5': 43094955}
         }]

# wb = xl.Workbook()
# ws = wb.active
# ws.title = 'Logs'
#
# for years_in_array, months_in_array in mydict.items():
#     ws.append([years_in_array])
#     for subMonths, total_calls in months_in_array.items():
#         ws.append([subMonths])
#         for call_labels, call_counts in total_calls.items():
#             ws.append([call_labels, call_counts])
#         ws.append([""])
# wb.save('C:/Users/a058943/Desktop/CALLS/Log Test.xlsx')

# Creating Data Frame

# d = {'Apple':{'Weight':12,'Colour':'red'},
#      'Banana':{'Weight':11,'Colour':'yellow','Bunched':1}
#     }
#
# df = pd.DataFrame.from_dict(mydict)
# df.to_csv('fruits.csv')
# # Creating two columns
# for years_in_array, months_in_array in mydict.items():
#     df['Year'] = mydict
#     for subMonths, total_calls in months_in_array.items():
#         df['Information'] = months_in_array[0:4]
#         for call_labels, call_counts in total_calls.items():
#             df['Calls'] = call_labels, call_counts

# Checking if file already exists

final_df = pd.DataFrame()

for id2 in range(0, len(data)):
    df = pd.DataFrame.from_dict(data[id2])
    final_df = pd.concat([final_df, df], axis=1)

print(final_df)

final_df.to_excel('data.xlsx')
print("Done")