require 'rubyXL'
require 'pry'
workbook = RubyXL::Workbook.new
#worksheet = workbook.add_worksheet('Sheet1')
worksheet = workbook[0] #Use first worksheet

worksheet.insert_row(0)
worksheet.insert_column(0)

worksheet.change_row_fill(0, '0bc53d')            # Sets first row to have fill #0ba53d
worksheet.change_column_fill(0, '0da54d')  

8.times.each_with_index do |item, row|
  #puts "current_index: #{index}"
  8.times.each_with_index do |item, column|
  	#puts "current_index: #{index}"
	  #worksheet.add_cell(row, column, '', 'A1').change_font_bold(true)
	  x=Random.rand(6)
	  puts "trial #{x}"
	  if(x==0)
	  	worksheet.add_cell(row, column, '', "SUM(#{row},#{column})").change_font_italics(true)
	  elsif(x==1)
	  	worksheet.add_cell(row, column, "#{row}test").change_font_name('Courier')
	  	worksheet.sheet_data[row][column].change_font_size(16)
	  elsif(x==2)
	  	worksheet.add_cell(row, column, "A1").change_horizontal_alignment('right')
	  	worksheet.sheet_data[row][column].change_font_underline(true)
	  elsif(x==3)
	  	worksheet.add_cell(row, column, "test#{column}").change_fill("0#{row}#{column}a3d")
	  else
	  	worksheet.add_cell(row, column, 'hey there').change_font_bold(true)
	  end
  	
  end
end
worksheet.merge_cells(0,0,0,1)
worksheet.merge_cells(1,4,2,4)

def cell_info(cell)
  puts cell.is_struckthrough 
  puts cell.font_size
  puts cell.font_color
  puts cell.fill_color
  puts cell.horizontal_alignment
  puts cell.vertical_alignment
  puts cell.get_border(:top)
  puts cell.get_border_color(:top)
  cell.text_rotation
end

#worksheet.add_cell(0, 0, '', 'SUM(1,1)')
workbook.write("file.xlsx")
binding.pry
