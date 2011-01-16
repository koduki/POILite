require 'java'
require 'jarlib/poi-3.7-20101029.jar'
require 'jarlib/poi-ooxml-3.7-20101029.jar'
require 'jarlib/poi-scratchpad-3.7-20101029.jar'

include_class 'java.io.FileInputStream'
include_class 'org.apache.poi.ss.usermodel.Sheet'
include_class 'org.apache.poi.ss.usermodel.Workbook'
include_class 'org.apache.poi.ss.usermodel.WorkbookFactory'

module POILite
  class WorkBook
    attr_reader :sheets
    def initialize poibook
      @poibook = poibook
      @sheets = @poibook.sheets.map{|sheet| WorkSheet.new sheet }
    end
  end

  class WorkSheets
    def initialize poibook
      @poibook = poibook
    end

    def [] index
      POILite::WorkSheet.new @poibook.getSheetAt(index)  
    end

    def size
      @poibook.sheets.size
    end

    include Enumerable
    def each &block
      @poibook.sheets.each {|sheet| block.call sheet }
    end

  end

  class WorkSheet
    attr_reader :rows

    def initialize poisheet
      @poisheet = poisheet 
      @rows = Rows.new @poisheet
    end

    def cells row_index, column_index
      @poisheet.getRow(row_index).getCell(column_index)  
    end

    def first_row_num 
      @poisheet.getFirstRowNum
    end
    def first_row
      @rows[first_row_num]
    end

    def last_row_num
      @poisheet.getLastRowNum
    end
    def last_row
      @rows[last_row_num]
    end

    def used_range
      used_rows = (first_row_num..last_row_num).map{ |i| @rows[i] }
      min_cell_num = used_rows.map{|r| r.first_cell_num }.min
      max_cell_num = used_rows.map{|r| r.last_cell_num }.max

      used_range = used_rows.map do |row|
        (min_cell_num..max_cell_num).map do |i|
          row.cells[i]  
        end
      end
    end
  end

  class Rows
    def initialize poisheet
      @poisheet = poisheet
    end

    def [](index)
      Row.new @poisheet.getRow(index)
    end
  end

  class Row
    attr_reader :cells

    def initialize poirow
      @poirow = poirow
      @cells = Cells.new @poirow
    end

    def first_cell_num 
      @poirow.getFirstCellNum
    end
    def first_cell
      @cells[first_cell_num]
    end

    def last_cell_num
      @poirow.getLastCellNum
    end
    def last_cell
      @cells[last_cell_num]
    end
  end

  class Cells
    def initialize poirow
      @poirow = poirow
    end

    def [](index)
      @poirow.getCell index
    end
  end
end

def open filename, &block
  input = FileInputStream.new filename
  poibook = WorkbookFactory.create(input)

  begin
    block.call POILite::WorkBook.new poibook
  ensure
    #poibook.close
    #input.close
  end
end

open("testcase.xls") do |book|
sheet1 = book.sheets[0]

row = sheet1.rows[0]
puts row.cells[0].to_s


p sheet1.first_row.last_cell
p sheet1.last_row.first_cell
puts sheet1.used_range.map{ |row|
  row.map{|cell| (cell != nil) ? cell.to_s : "" }.join(",") 
}.join("\n")

cell = sheet1.cells(0, 0)
p cell.getCellType
end
