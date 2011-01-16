require 'poilite.rb'

POILite::Excel::open("testcase.xls") do |book|
  sheet1 = book.sheets[0]
  
  row = sheet1.rows[0]
  puts row.cells[0]
  
  p sheet1.first_row.last_cell
  p sheet1.last_row.first_cell
  p sheet1.cells(0, 7)

  puts sheet1.used_range.map{ |row|
    row.map{|cell| (cell != nil) ? cell.to_s : "" }.join(",") 
  }.join("\n")

end
