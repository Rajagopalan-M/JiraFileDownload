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
    @b = Watir::Browser.new :chrome, options: {prefs: prefs}, args: ['user-data-dir=C:\ChromeProfile\ChromeProfile']
    @b.window.maximize
    @b.goto 'https://jira.sapiens.com/secure/Dashboard.jspa'
    if @b.text_field(name: 'os_username').present?
      @b.text_field(name: 'os_username').set 'rajagopalan.m'
      @b.text_field(name: 'os_password').set 'ceasar@123'
      @b.button(name: 'login').click
    end
    @b.element(link: 'Issues').click
    @b.element(link: 'Search for issues').click
    @b.element(link: 'Advanced').click unless @b.element(link: 'Basic').present?
    @b.button(id: 'layout-switcher-button').click
    @b.link(text: 'List View').click
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

  def newFetchAndWrite
    @inputSheet.each.with_index do |(inputKey, inputValues), index|
      outputExcel = WriteExcel.new
      inputValues.each do |inputValue|
        @b.textarea(id: 'advanced-search').set inputValue[3]
        @b.button(xpath: "//div[@class='search-options-container']/button").click
        sleep 2
        totalResults = @b.span(class: ["results-count-total", "results-count-link"]).text.to_i
        query = ""
        queryArray = [inputValue[3]]
        project = (inputValue[3].split(/(project\s+=\s+\w+)/).drop(1))[0].split('=').last.strip
        moreResultsFlag = false
        if totalResults > 1000
          queryArray = []
          moreResultsFlag = true
          @b.span(title: 'Sort By Key').click
          lastKeyValue = @b.a(xpath: "(//a[@class='issue-link'])[1]").text.split('-').last.to_i
          timesValues = (lastKeyValue.to_f / 1000.00).ceil
          startVal = 0
          endVal = 1000
          timesValues.times do |i|
            if startVal.eql? 0
              loop do
                query = inputValue[3].split(/(project\s+=\s+\w+)/).drop(1).join(" AND key <= #{project}-#{endVal}")
                @b.textarea(id: 'advanced-search').set query, :return
                sleep 1
                break unless @b.div(class: ["aui-message", "aui-message-error"]).present?
                endVal += 1 if @b.div(class: ["aui-message", "aui-message-error"]).present?
              end
            elsif i.eql? (timesValues - 1)
              query = inputValue[3].split(/(project\s+=\s+\w+)/).drop(1).join(" AND key > #{project}-#{startVal} AND key <= #{project}-#{lastKeyValue}")
            else
              loop do
                query = inputValue[3].split(/(project\s+=\s+\w+)/).drop(1).join(" AND key > #{project}-#{startVal} AND key <= #{project}-#{endVal}")
                @b.textarea(id: 'advanced-search').set query, :return
                sleep 1
                break unless @b.div(class: ["aui-message", "aui-message-error"]).present?
                endVal += 1 if @b.div(class: ["aui-message", "aui-message-error"]).present?
              end
            end
            startVal = endVal
            if (endVal.modulo 1000).eql? 0
              endVal += 1000
            else
              endVal = (endVal / 1000.0).ceil * 1000
            end
            queryArray << query
          end #timesValues
        end

        csvTable = []
        queryArray.each.with_index do |query, queryArrayIndex|
          @b.textarea(id: 'advanced-search').set query, :tab
          @b.button(xpath: "//div[@class='search-options-container']/button").click
          #Select columns for output
          @b.button(title: 'Columns').click
          @reference[[inputValue[0], inputValue[2]]].each do |jiraScreenName|
            @b.text_field(id: 'user-column-sparkler-input').set jiraScreenName[3]
            @b.checkbox(xpath: "//ul[@class='aui-list-section aui-last']/li/label/em[normalize-space()='#{jiraScreenName[3]}' and not(preceding-sibling::span)]/preceding-sibling::input").set
          end unless @reference[[inputValue[0], inputValue[2]]].nil?
          sleep 0.5
          @b.button(value: 'Done').click

          #output in csv format
          unless @b.div(class: ["aui-message", "error"]).present?
            sleep 2
            @b.scroll.to :top
            @b.span(text: 'Export').click
            sleep 1
            @b.link(id: 'currentCsvFields').click
            @b.button(id: 'csv-export-dialog-export-button').click
            sleep 2
          end

          #read csv and put it in outputExcel.
          table = CSV.read(Dir["Downloads/*.csv"].sort { |a, b| File.mtime(a) <=> File.mtime(b) }.last)
          table.drop(1) if moreResultsFlag and queryArrayIndex > 0

          sindex = table[0].index('Sprint')
          count = table[0].count('Sprint')
          processedCSVTable = []
          table.map do |row|
            newRow = []
            newRow << row[sindex, count].uniq.join(',').chomp(',')
            row.each.with_index { |val, ind| newRow << val unless (sindex..(sindex + (count - 1))).include? ind }
            processedCSVTable << newRow
          end

          csvTable = csvTable + processedCSVTable
        end #queryArray

        excelSheet = outputExcel.createSheet inputValue[2]
        outputTable = []
        @reference[[inputValue[0], inputValue[2]]].each do |columnName|
          csvTable.transpose.each do |transposeRow|
            if transposeRow[0].eql? columnName[2]
              transposeRow[0] = columnName[4]
              outputTable << transposeRow
            end
          end
        end
        finalOutputTable = outputTable.transpose
        excelSheet.enterTheData(finalOutputTable, [finalOutputTable[0].index("Summary"), finalOutputTable[0].index("Sprint"), 50])

      rescue => e
        puts e.message
        @b.screenshot.save ("screenshot.png")
      end
      outputExcel.write(File.expand_path("Output/#{(inputKey[0] + " " + inputKey[1] + " " + "File" + " ").to_s + Date.today.strftime('%Y%m%d').gsub('-', '_')}.xlsx"))
    end
  end

  def fetchAndWrite
    @inputSheet.each.with_index(1) do |(k, arrs), index|
      #Read the Epic Key and Epic Name to form the Hash : (Epic Key=>Epic Name)
      puts k.first
      @b.textarea(id: 'advanced-search').set "project = #{k.first[0..2]} AND issuetype = Epic", :return
      FileUtils.rm_rf Dir.glob(File.expand_path("Downloads/*"))
      @b.span(text: 'Export').click
      @b.link(id: 'currentCsvFields').click
      @b.button(id: 'csv-export-dialog-export-button').click
      sleep 1
      @b.wait_while do
        @b.wait_until { Dir.glob(File.expand_path("Downloads/*")).count > 0 }
        filename = Dir.glob(File.expand_path("Downloads\\*.*")).last
        filename.include? 'crdownload'
      end
      tablee = CSV.read(Dir.glob(File.expand_path("Downloads/*")).first)
      table = tablee.transpose.select { |x| x[0].eql? 'Issue key' or x[0].eql? 'Custom field (Epic Name)' }
      issue_key = table[0].zip(tablee[1]).to_h
      ####
      workBook = WriteExcel.new
      arrs.each do |arr|
        sheet = workBook.createSheet(arr[2])
        flag = false
        # //span[@class='results-count-total results-count-link']

        table = []
        query = nil
        loop do

          if flag.eql? false
            @b.textarea(id: 'advanced-search').set arr[3], :return
          else
            @b.textarea(id: 'advanced-search').set query, :return
          end

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
            @b.button(id: 'csv-export-dialog-export-button').click
            sleep 2
            filename = nil
            @b.wait_while do
              filename = Dir.glob(File.expand_path("Downloads\\*.*")).last
              filename.include? 'crdownload'
            end
            csvTable = CSV.read(filename)
            p csvTable.transpose
            table << @reference[[arr[0], arr[2]]]
                         .map { |row| csvTable.transpose.select { |tableRow| tableRow[0].eql? row[2] }.pop }
                         .reject { |x| x.to_s.empty? }
                         .transpose unless @reference[[arr[0], arr[2]]].nil?
            countOfTotalRows = @b.span(class: ["results-count-total", "results-count-link"]).text.to_i
            if countOfTotalRows < 1000
              break
            else
              key = table[0][0].index('Issue key')
              row = table.flatten(1).transpose[key]
              query = arr[3].split(/(project\s+=\s+\w+)/).drop(1).join(" and Key > #{row.last} ")
              flag = true
            end
          end
        end
        table = table.flatten(1)
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
    .newFetchAndWrite

