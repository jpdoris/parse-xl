import xlrd

workbook = xlrd.open_workbook('SampleData.xls')
sheet = workbook.sheet_by_name('SalesOrders')
itemtofind = "Pen"
itemcount = 0
itemtotal = 0

# loop through header cells to identify 'Item', 'Units', and 'Total' col indices
for col_index in xrange(1, sheet.ncols):
  if (sheet.cell(0,col_index).value) == 'Item':
    itemcol = col_index
  if (sheet.cell(0,col_index).value) == 'Units':
    unitcol = col_index
  if (sheet.cell(0,col_index).value) == 'Total':
    totalcol = col_index

# loop through data cells to find all pens (or whichever item we specified in
# 'itemtofind') and get the count, and the total, in increments
for row_index in xrange(1, sheet.nrows):
  if (sheet.cell(row_index,itemcol).value == itemtofind):
    try:
      if (sheet.cell(row_index,unitcol).value):
        itemcount += int(sheet.cell(row_index,unitcol).value)
      if (sheet.cell(row_index,totalcol).value):
        itemtotal += float(sheet.cell(row_index,totalcol).value)
    except ValueError:
      pass

# print the results
print "Total " + itemtofind + "s ordered: " + str(itemcount)
print "Total spent on " + itemtofind + "s: $" + str(itemtotal)
