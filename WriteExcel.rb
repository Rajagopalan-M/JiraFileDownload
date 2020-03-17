#Creator : Rajagopalan M
#Date    : 19-09-2019
########################

# project = AIP AND issuetype = Epic

class WriteExcel
  def initialize
    require 'rubyXL'
    @workBook = RubyXL::Workbook.new
  end

  def createSheet sheetname
    @worksheet = @workBook['Sheet1']
    unless @worksheet.nil?
      @worksheet.sheet_name = sheetname
    else
      @worksheet = @workBook.add_worksheet(sheetname)
    end
    self
  end

  def enterTheData arr, wrap = nil
    filter = RubyXL::AutoFilter.new
    range = RubyXL::Reference.new(0, arr.count - 1, 0, arr.first.count - 1)
    filter.ref = range

    @worksheet.auto_filter = filter

    arr.transpose.map { |x| x.map { |y| y.to_s.strip.chars.count }.max }.each_with_index do |length, index|
      length = 35 if length > 50
      @worksheet.change_column_width(index, length + 1)
    end

    #@worksheet.change_column_width(wrap[0],wrap[1]) unless wrap.nil?

    arr.each_with_index do |row, rowIndex|
      row.each_with_index do |data, colIndex|
        @worksheet.add_cell(rowIndex, colIndex, data)
        if rowIndex.eql? 0
          @worksheet.sheet_data[rowIndex][colIndex].change_fill('2f5597')
          @worksheet.sheet_data[rowIndex][colIndex].change_font_color('ffffff')
          @worksheet.sheet_data[rowIndex][colIndex].change_font_bold(true)
        end
        unless wrap.nil?
          @worksheet[rowIndex][colIndex].change_text_wrap(true) if colIndex.eql? wrap[0] or colIndex.eql? wrap[1]
        end
        @worksheet.sheet_data[rowIndex][colIndex].change_border(:top, 'thin')
        @worksheet.sheet_data[rowIndex][colIndex].change_border(:bottom, 'thin')
        @worksheet.sheet_data[rowIndex][colIndex].change_border(:right, 'thin')
        @worksheet.sheet_data[rowIndex][colIndex].change_border(:left, 'thin')
      end
    end
    self
  end

  def write filename
    @workBook.write(filename)
  end
end
