from pyspark import SparkConf, SparkContext
import re
import xlsxwriter

# if another SparkContext was running
sc.stop()

# set spark context
conf = SparkConf().setMaster("local")
sc = SparkContext(conf=conf)

# text file
txt_input = sc.textFile("daftar5.txt")

# map-reduce
char_counts = txt_input.flatMap(lambda x: [each[0] if each != '' else 'Empty' for each in re.split(' |\t\t|\t|\\u200f', x)]).map(lambda char: char).map(lambda c: (c, 1)).reduceByKey(lambda v1, v2: v1 + v2)

# collect
lst = char_counts.collect()

# create .xlsx file
workbook = xlsxwriter.Workbook('first_char_count_output.xlsx')

# add new worksheet
worksheet = workbook.add_worksheet()

# characters in A col
worksheet.write(str('A' + str(1)), 'Characters')

# counts in C col
worksheet.write(str('C' + str(1)), 'Counts')

# store in the .xlsx file
counter = 2
for i in lst:
    if i[0] != 'Empty':
      # characters in A2 - An col
      worksheet.write(str('A' + str(counter)), i[0])

      # counts in C2 - Cn col
      worksheet.write(str('C' + str(counter)), i[1])

      counter += 1

# close workbook
workbook.close()