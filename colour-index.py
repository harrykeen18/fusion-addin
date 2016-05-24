from xlrd import open_workbook
import math
import operator

OUTFILE = '/Users/harry/Library/Application Support/Autodesk/Autodesk Fusion 360/API/Scripts/tester/assets/acad-ref-dict.txt'

def importer():

  file_loc = '/Users/harry/Desktop/LAYERCOLOURS - new.xlsx'

  wb = open_workbook(file_loc)

  sheet = wb.sheet_by_index(0)

  actual_dict = {}
  for colnum in range(1, sheet.ncols):
    col_values_list = []
    for rownum in range(1, sheet.nrows):
      col_values_list.append(eval(sheet.cell_value(rownum, colnum)))

    actual_dict[sheet.cell_value(0, colnum)] = col_values_list

  sheet = wb.sheet_by_index(2)

  ref_dict = {}
  for colnum in range(0, sheet.ncols):
    col_values_list = []
    for rownum in range(1, sheet.nrows):
      col_values_list.append(sheet.cell_value(rownum, colnum))

    ref_dict[sheet.cell_value(0, colnum)] = col_values_list

  n = 0 

  for each_key in actual_dict.keys():
    for ind, array in enumerate(actual_dict[each_key]):

      red_diff_list = []
      blue_diff_list = []
      green_diff_list = []

      for i, acad_index in enumerate(ref_dict['index']):
        red_diff_list.append(math.sqrt(math.pow((ref_dict['r'][i] - array[0]), 2)))
        green_diff_list.append(math.sqrt(math.pow((ref_dict['g'][i] - array[1]), 2)))
        blue_diff_list.append(math.sqrt(math.pow((ref_dict['b'][i] - array[2]), 2)))

        # ref_dict['']

        # print acad_index

      indexed_diff_dict = {}
      for i in range(0, len(red_diff_list)):
        diff = (red_diff_list[i] + blue_diff_list[i] + green_diff_list[i])
        indexed_diff_dict[i] = diff
      
      sorted_indexed_dict = sorted(indexed_diff_dict.items(), key=operator.itemgetter(1))

      nearest_val = sorted_indexed_dict[0]

      a = array
      a.append(nearest_val[0])
      #print a
      actual_dict[each_key][ind] = a

  with open(OUTFILE, "w"):
    pass
  
  acad_ref_txt = open(OUTFILE, 'w')

  for each_key in actual_dict.keys():
    acad_ref_txt.write("%s\n" % str(each_key))
    for array in actual_dict[each_key]:
      acad_ref_txt.write("%s\n" % str(array))

  acad_ref_txt.close()


if __name__ == '__main__':
  importer()

