#Creator : Rajagopalan M
#Date    : 19-Sep-2019
########################
require_relative 'WriteExcel'
class Jira
  def initialize
    require 'watir'
    require 'rubyXL'
    require 'creek'
    require 'csv'
    ["Downloads", "Output"].each do |folderName|
      FileUtils.rm_rf File.expand_path(folderName)
      FileUtils.mkdir File.expand_path(folderName)
    end
    prefs = {
        download: {
            prompt_for_download: false,
            default_directory: File.expand_path("Downloads").gsub('/', '\\')
        }
    }
    @b = Watir::Browser.new :chrome, options: {prefs: prefs}, args: ['user-data-dir=C:\RajagopalanM']
    @b.window.maximize
    @b.goto 'https://jira.sapiens.com/secure/Dashboard.jspa'
    # @b.text_field(name: 'os_username').set 'rajagopalan.m'
    # @b.text_field(name: 'os_password').set 'gopalan@456'
    # @b.button(name: 'login').click
    @b.element(link: 'Issues').click
    @b.element(link: 'Search for issues').click
    @b.element(link: 'Advanced').click unless @b.element(link: 'Basic').present?
    @workbook = Creek::Book.new 'Input.xlsx'
  end

  def readReferenceSheet
    @referenceSheet = @workbook.sheets[1].rows.map(&:values)

    @referenceSheet.each_cons(2) do |first, last|
      last[0] ||= first[0] if (last[0].nil? or last[0].to_s.empty?)
      last[1] ||= first[1] if (last[1].nil? or last[1].to_s.empty?)
    end
    @reference = @referenceSheet.drop(1).group_by { |x| [x[0], x[1]] }
    self
  end

  def readInputSheet
    #Read the sheet and group the value by company name and Daily/Weekly
    @inputSheet = @workbook.sheets[0].rows.map(&:values).drop(1).group_by { |x| [x[0], x[1]] }
    self
  end

  def fetchAndWrite
    @inputSheet.each.with_index(1) do |(k, arrs), index|
      #Read the Epic Key and Epic Name to form the Hash : (Epic Key=>Epic Name)
      @b.textarea(id: 'advanced-search').set "project = #{k.first[0..3]} AND issuetype = Epic", :return
      FileUtils.rm_rf Dir.glob(File.expand_path("Downloads/*"))
      @b.span(text: 'Export').click
      @b.link(id: 'currentCsvFields').click
      sleep 1
      @b.wait_while do
        @b.wait_until { Dir.glob(File.expand_path("Downloads/*")).count > 0 }
        filename = Dir.glob(File.expand_path("Downloads\\*.*")).last
        filename.include? 'crdownload'
      end
      table = CSV.read(Dir.glob(File.expand_path("Downloads/*")).first)
      table = table.transpose.select { |x| x[0].eql? 'Issue key' or x[0].eql? 'Custom field (Epic Name)' }
      issue_key = table[0].zip(table[1]).to_h
      ####
      workBook = WriteExcel.new
      arrs.each do |arr|
        sheet = workBook.createSheet(arr[2])
        @b.textarea(id: 'advanced-search').set arr[3], :return
        @b.button(title: 'Columns').click
        @reference[[arr[0], arr[2]]].each do |field|
          @b.text_field(id: 'user-column-sparkler-input').set field[4]
          @b.checkbox(xpath: "//ul[@class='aui-list-section aui-last']/li/label/em[normalize-space()='#{field[4]}' and not(preceding-sibling::span)]/preceding-sibling::input").set
        end unless @reference[[arr[0], arr[2]]].nil?
        @b.button(value: 'Done').click
        $stdout.sync = true
        unless @b.div(class: ["aui-message", "error"]).present?
          sleep 0.5
          @b.scroll.to :top
          @b.span(text: 'Export').click
          @b.link(id: 'currentCsvFields').click
          sleep 2
          filename = nil
          @b.wait_while do
            filename = Dir.glob(File.expand_path("Downloads\\*.*")).last
            filename.include? 'crdownload'
          end
          table = CSV.read(filename)
          table = @reference[[arr[0], arr[2]]].map { |row| table.transpose.select { |tableRow| tableRow[0].eql? row[2] }.pop }.reject { |x| x.to_s.empty? }.transpose unless @reference[[arr[0], arr[2]]].nil?

          resultColumn = @reference[[arr[0], arr[2]]].map { |x| x[4] } unless @reference[[arr[0], arr[2]]].nil?

          #Form the reference hash(this is used for replacing the csv headers with our headers)
          h = @referenceSheet.drop(1).each_with_object({}) do |value, h|
            h[value[2]] = value[4]
          end
          # Changing the table header with the our header.
          table[0].each_with_index do |v, index|
            table[0][index] = h[v]
          end
          #Replace the Epic Key with Epic Name
          ind = nil
          table.each_with_index do |row, index|
            if index.eql? 0
              ind = row.index("Epic Link")
              # row[4] = 'Epic Name'
              next
            end
            row[ind] = issue_key[row[ind]]
          end
          table = table.transpose.select { |x| resultColumn.include? x[0].to_s.strip }.transpose
        end
        sheet.enterTheData(table, [table[0].index("Summary"), 50]) unless table.nil?
      end

      workBook.write(File.expand_path("Output/#{(k[0] + " " + k[1] + " " + "File" + " ").to_s + Date.today.strftime('%Y%m%d').gsub('-', '_')}.xlsx"))
    end
    FileUtils.rm_rf File.expand_path("Downloads")
  end
end

Jira.new
    .readReferenceSheet
    .readInputSheet
    .fetchAndWrite

