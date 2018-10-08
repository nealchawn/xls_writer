require 'rubyXL'
require 'pry'
workbook = RubyXL::Workbook.new
#worksheet = workbook.add_worksheet('Sheet1')
worksheet = workbook[0] #Use first worksheet

#make 7x5 rows and columns
7.times.each_with_index do |item, row|
  5.times.each_with_index do |item, column|
    worksheet.add_cell(row, column ,"")
  end
end

#merge_cells(start_row,col,end_row,col)
#Title line
worksheet.merge_cells(0,0,0,4)
worksheet[0][0].change_contents("this is an example of 5 merged cells")


#Column headers
5.times.each_with_index do |item, column|
  worksheet[2][column].change_contents("Column #{column}")
  worksheet[2][column].change_font_bold(true)
  worksheet[2][column].change_fill("eeeeee")
end

#Row data
4.times.each_with_index do |item, row|
  5.times.each_with_index do |item, column|
    worksheet[row+3][column].change_contents("Test Data: #{row} #{column}")
  end
end

workbook.write("file.xlsx")
binding.pry
