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
    @b.text_field(name: 'os_username').set 'rajagopalan.m'
    @b.text_field(name: 'os_password').set 'gopalan@456'
    @b.button(name: 'login').click
    @b.element(link: 'Issues').click
    @b.element(link: 'Search for issues').click
    @b.element(link: 'Advanced').click unless @b.element(link: 'Basic').present?
    @workbook = Creek::Book.new 'Input.xlsx'
  end

  def readReferenceSheet
    @referenceSheet = @workbook.sheets[1].rows.map(&:values)
    customerName = nil
    tab = nil
    # @referenceSheet.each do |arr|
    #   if arr[0].nil?
    #     arr[0] = customerName
    #   else
    #     customerName = arr[0]
    #   end
    #   if arr[1].nil?
    #     arr[1] = tab
    #   else
    #     tab = arr[1]
    #   end
    # end
    @referenceSheet.each_cons(2) do |first, last|
      last[0] ||= first[0] if (last[0].nil? or last[0].to_s.empty?)
      last[1] ||= first[1] if (last[1].nil? or last[1].to_s.empty?)
    end
    @reference = @referenceSheet.drop(1).group_by { |x| [x[0], x[1]] }
    self
  end

  def readInputSheet
    @inputSheet = @workbook.sheets[0].rows.map(&:values).drop(1).group_by { |x| [x[0], x[1]] }
    self
  end

  def fetchAndWrite
    @inputSheet.each.with_index(1) do |(k, arrs), index|

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

          h = @referenceSheet.drop(1).each_with_object({}) do |value, h|
            h[value[2]] = value[4]
          end
          table[0].each_with_index do |v, index|
            table[0][index] = h[v]
          end
        end
        sheet.enterTheData(table, [table[0].index("Summary"), 50]) unless table.nil?
      end

      workBook.write(File.expand_path("Output/#{(k[0] + " " + k[1] + " " + "File" + " ").to_s + Date.today.strftime('%Y%m%d').gsub('-', '_')}.xlsx"))
    end
    # FileUtils.rm_rf File.expand_path("Downloads")
  end
end

Jira.new
    .readReferenceSheet
    .readInputSheet
    .fetchAndWrite