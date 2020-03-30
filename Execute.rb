#Creator : Rajagopalan M
#Edited : Chetan Kondawle (Methods : restFetchAndWrite)
#Date    : 19-Sep-2019
########################
require_relative 'WriteExcel'
class Jira
  def initialize userName, password
    require 'rubyXL'
    require 'creek'
    require 'json'
    require 'base64'
    require 'rest-client'
    FileUtils.mkdir_p 'Logs'
    FileUtils.mkdir_p "JiraOutput/#{Time.now.strftime "%d %^b %Y"}"
    @fw = File.open("Logs/#{Time.now.strftime "%Y-%m-%d_%H-%M-%S"}.txt", 'w')
    @programStartTime = Time.now

    @workbook = Creek::Book.new 'Input.xlsx'
    @url = 'https://jira.sapiens.com/rest'
    @userName = userName
    @password = password
    @epicHash = Hash.new
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

  def restFetchAndWrite
    @inputSheet.each.with_index do |(inputKey, inputValues), index|
      if inputKey[1].downcase.eql? 'weekly'
        next unless Date.today.strftime("%A").downcase.eql? 'thursday'
      end
      outputExcel = WriteExcel.new
      inputValues.each do |inputValue|
        excelSheet = outputExcel.createSheet inputValue[2]

        #Forming Query and Fetching total results
        query = inputValue[3].gsub(/[\s=&,]/, ' ' => '%20', '=' => '%3D', ',' => '%2C', '&' => '%26')
        filterColumns = @reference[[inputValue[0], inputValue[2]]].map { |jsonFieldName| jsonFieldName[5] }.join(',')
        auth = "Basic #{Base64.encode64("#{@userName}:#{@password}")}"
        response = RestClient.get "#{@url}/api/2/search?jql=#{query}&fields=#{filterColumns}&maxResults=0", {Authorization: auth}
        totalResults = (JSON.parse(response.body))['total'].to_i
        puts "Project : #{inputValue[0]} | Tab : #{inputValue[2]} | Total Results : #{totalResults}"
        @fw.puts "Project : #{inputValue[0]} | Tab : #{inputValue[2]} | Total Results : #{totalResults}"
        #Forming Query and Fetching total results

        if totalResults > 0
          dataArray = [@reference[[inputValue[0], inputValue[2]]].map { |excelColumnName| excelColumnName[4] }]
          refArr = @reference[[inputValue[0], inputValue[2]]].map { |refColumnName| refColumnName[6] }

          #Fetching and processing data
          startIndex = 0
          (startIndex..totalResults).step(1000) do |n|
            response = RestClient.get "#{@url}/api/2/search?jql=#{query}&fields=#{filterColumns}&maxResults=1000&startAt=#{n}", {Authorization: auth}
            parsedJson = JSON.parse(response.body)
            #retrieving data from JSON Response
            parsedJson['issues'].each.with_index do |issuesHash, index|
              tempArr = []
              refArr.each.with_index do |refHashQuery, index|
                tempIssue = issuesHash
                refHashQuery = refHashQuery.split(',')
                refHashQuery.count.times do |i|
                  tempIssue = fetch(tempIssue, refHashQuery[i]) unless tempIssue.nil?
                end
                if refArr[index].eql? 'fields,customfield_10004'
                  tempIssue = tempIssue.map { |x| (x.split /,name=/).last.split(',').first }.join(',') unless tempIssue.nil?
                elsif refArr[index].eql? 'fields,customfield_10000'
                  if @epicHash.include? tempIssue.strip
                    tempIssue = @epicHash[tempIssue.strip]
                  else
                    @epicHash[tempIssue.strip] = (JSON.parse((RestClient.get "#{@url}/agile/1.0/epic/#{tempIssue}", {Authorization: auth}).body))['name']
                    tempIssue = @epicHash[tempIssue.strip]
                  end unless tempIssue.nil?
                elsif refArr[index].eql? 'fields,assignee,displayName'
                  tempIssue = 'Unassigned' if tempIssue.nil?
                elsif refArr[index].eql? 'fields,resolution,name'
                  tempIssue = 'Unresolved' if tempIssue.nil?
                elsif refArr[index].eql? 'fields,created'
                  tempIssue = (Time.parse(tempIssue) + 7200).strftime('%m/%d/%Y %H:%M')
                end
                tempArr << tempIssue
              end
              tempArr = tempArr.map { |x| x.nil? ? "" : x }
              dataArray << tempArr
            end
          end
          #Fetching and processing data
          excelSheet.enterTheData(dataArray, [dataArray[0].index("Summary"), dataArray[0].index("Sprint"), 50])
        else
          excelSheet.zeroResult
        end
      rescue Exception, RestClient::ExceptionWithResponse => e
        e.message = "Wrong Username & Password" if e.instance_of? RestClient::Unauthorized
        e.message = (JSON.parse(e.http_body))['errorMessages'].join(" ") if e.instance_of? RestClient::BadRequest
        @fw.puts "Project : #{inputValue[0]} | Tab : #{inputValue[2]} | Error : " + e.message
        puts "Project : #{inputValue[0]} | Tab : #{inputValue[2]} | Error : " + e.message
        break if e.instance_of? RestClient::Unauthorized
      end
      outputExcel.write(File.expand_path("JiraOutput/#{Time.now.strftime "%d %^b %Y"}/#{(inputKey[0] + " " + inputKey[1] + " " + "File" + " ").to_s + Date.today.strftime('%Y%m%d').gsub('-', '_')}.xlsx"))
    end
  ensure
    @programEndTime = Time.now
    seconds = (@programEndTime - @programStartTime).to_i
    min = (seconds.to_f / 60).to_s.split('.').first.to_i
    seconds = seconds - (min * 60)
    @fw.puts "=================================="
    @fw.puts "Program Ended : Time : #{min}m.#{seconds}s"
    @fw.close
  end

  private

  def fetch (hash, key)
    hash[key]
  end

end

puts "Enter username for Jira :"
user = gets.chomp
puts "Enter password for Jira :"
pass = gets.chomp

Jira.new(user, pass)
    .readReferenceSheet
    .readInputSheet
    .restFetchAndWrite