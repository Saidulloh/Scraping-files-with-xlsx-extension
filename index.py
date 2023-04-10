# from openpyxl import load_workbook

########## Code for first table ##########

# wb = load_workbook(filename='index_1.xlsx')
# sheet = wb['Люстра или подвес (под лампу)']

# i = 0
# s = 1
# some = 77
# number = 77
# for row in sheet.iter_rows():
#     for cell in row:
#         i += 1
#         if i <= number:
#             with open(f'index_1/data_{s}', 'a+', encoding='utf-8') as file:
#                 file.write(f'{cell.value}\n')
#                 file.close()
#         else:
#             s += 1
#             number += some

########## End code for first table ##########


########## Code for second table ##########

# wb = load_workbook(filename='index_2.xlsx')

# sheet = wb['Люстра или подвес светодиодные ']
# i = 0
# s = 1
# some = 72
# number = 72
# for row in sheet.iter_rows():
#     for cell in row:
#         i += 1
#         print(i, cell.value)
#         if i <= number:
#             with open(f'index_2/data_{s}', 'a+', encoding='utf-8') as file:
#                 file.write(f'{cell.value}\n')
#                 file.close()
#         else:
#             s += 1
#             number += some

########## End code for second table ##########
