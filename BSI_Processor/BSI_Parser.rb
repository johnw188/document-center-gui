# This set of classes allows BSI feeder data to be extracted from the monthly

require 'win32ole'
require 'rubygems'
require 'mechanize'

class Source_data
 attr_reader :dataHash
 def initialize()
   @dataHash = {}
 end
 def print
   @dataHash.each_pair{|docName, valueArray|
     puts "#{docName}: #{valueArray.join("\n")} \n"
   }
 end
end

class BSI_Website_Parser
  def initialize
    @agent = WWW::Mechanize.new
    page = @agent.get 'http://www.bsi-global.com/en/My-BSI/My-Subscriptions/'
    @form = page.forms[0]
  end

  def searchForDocument(docTitle)
    @form.q = docTitle
    page = @agent.submit(@form)
    page = @agent.click page.links.find { |l| l.text =~ Regexp.new(docTitle)}
    pars = ""
    page.search("//div[@id=tab2]").each {|p|
      pars << p.inner_html
    }
    return pars
  end
end

class BSI_excel_data < Source_data
  attr_reader :dataHash, :columnValues
  def initialize(filepath)
    excel = WIN32OLE::new('excel.Application')
    workbook = excel.Workbooks.Open(filepath)
    worksheet = workbook.Worksheets(1)
    column = 'a'
    while worksheet.Range("#{column.succ}1")['Value']
      column.succ!
    end
    @dataHash = {}
    @columnValues = worksheet.Range("a1:#{column}1")['Value'].flatten
    line = '2'
    while worksheet.Range("a#{line}")['Value']
      infoArray = worksheet.Range("a#{line}:#{column}#{line}") ['Value'].flatten
      infoArray.push line.to_i
      @dataHash[worksheet.Range("a#{line}")['Value']] = infoArray
      line.succ!
    end
    excel.Quit
  end
  
 def getNewItems(olderBSIdata)
   newItemsHash = {}
   @dataHash.each_pair{|key, value|
     newItemsHash[key] = value unless olderBSIdata.dataHash.has_key?(key)
   }
   newItemsHash
 end
 def fixDates
   column = @columnValues.rindex("Publication date")
   @dataHash.each_value {|value|
     newDate = value[column].split("/")
     value[column] = "#{newDate[1]}/#{newDate[0]}/#{newDate[2]}"
   }
 end
 def prepareOutput
   outputArray = []
   @dataHash.each_value {|value|
     documentIdentifier = value[0]
     splitName = documentIdentifier.split(":")
     invtid = splitName[0].gsub(" ","-")
     suffixid = "#{splitName[1]} EDITION"
     titleIdentifier = value[1]
     title1 = titleIdentifier[0,60] ? titleIdentifier[0,60].upcase : ""
     title2 = titleIdentifier[60,30] ? titleIdentifier [60,30].upcase : ""
     status = value[2]
     publicationDate = value[3]
     committee = value[4] ? value[4] : ""
     price = value[5] ? value[5].sub("\243","") : ""
     isbn = value[7]
     pages = value[8] ? value[8] : ""
     replaces = value[12] ? value[12] : ""
     replacedBy = value[14] ? value[14] : ""
     tempDataArray = [invtid, suffixid, documentIdentifier,  titleIdentifier, title1, title2, status, publicationDate, committee,  price, isbn, pages, replaces, replacedBy]
     outputArray << tempDataArray
   }
   outputArray
 end
end

def prepareOutput(dataHash)
   outputArray = []
   dataHash.each_value {|value|
     documentIdentifier = value[0]
     splitName = documentIdentifier.split(":")
     invtid = splitName[0].gsub(" ","-")
     suffixid = "#{splitName[1]} EDITION"
     titleIdentifier = value[1]
     title1 = titleIdentifier[0,60] ? titleIdentifier[0,60].upcase : ""
     title2 = titleIdentifier[60,30] ? titleIdentifier [60,30].upcase : ""
     status = value[2]
     publicationDate = value[3]
     committee = value[4] ? value[4] : ""
     price = value[5] ? value[5].sub("\243","") : ""
     isbn = value[7]
     pages = value[8] ? value[8] : ""
     replaces = value[12] ? value[12] : ""
     replacedBy = value[14] ? value[14] : ""
     tempDataArray = [invtid, suffixid, documentIdentifier,  titleIdentifier, title1, title2, status, publicationDate, committee,  price, isbn, pages, replaces, replacedBy]
     outputArray << tempDataArray
   }
   outputArray
 end

def writeToExcel(array, filename)
 excel = WIN32OLE::new('excel.Application')
 workbook = excel.Workbooks.Add
 worksheet = workbook.Worksheets(1)
 worksheet.select
 columns = array[0].size
 column = 'a'
 (columns - 1).times do column.succ! end
 row = array.size
 worksheet.Range("a1:#{column}#{row}")['Value'] = array
 workbook.SaveAs("#{Dir.pwd}/#{filename}")
 excel.Quit
end

def writeInfoPage(docNameArray, filename)
  bsiParser = BSI_Website_Parser.new
  htmlFile = File.new(filename, "wb")
  htmlFile.puts "<html>
  <head>
  <title>
  Latest BSI documents (#{docNameArray.size.to_s} items)
  </title>

  <style type=\"text/css\">
    div.contentZone table {
  border-collapse:collapse;
  clear:both;
  width:100%;
  }
  table.bibliography th, table.bibliography td {
  background:transparent url(/dotted_bg.gif) repeat-x scroll left bottom;
  padding:4px 0pt 6px;
  vertical-align:top;
  }
  table.bibliography th {
  color:#676767;
  text-align:left;
  }
  table.bibliography td {
  background:transparent url(/dotted_bg.gif) repeat-x scroll left bottom;
  padding:4px 0pt 6px;
  vertical-align:top;
  }
  table.bibliography * {
  border:0pt none;
  }
  table th {
  background:#C6C7C7 none repeat scroll 0%;
  border-top:0pt none;
  font-weight:normal;
  text-align:left;
  }
  table th, table td {
  border:2px solid #F1F1F1;
  color:#333333;
  padding:5px;
  }
  table.bibliography {
  margin:0.5em 0pt 1em;
  }
  table {
  border-collapse:collapse;
  width:auto;
  }
  table {
  font-size:1.1em;
  }

  html {
  font-family:Arial,Helvetica,Sans-Serif;
  font-size:62.5%;
  }
  </style>
  </head>
  <body>"
  docNameArray.each {|document|
    data = bsiParser.searchForDocument(document)
    startOfHTML = data.index("<table class=")
    htmlFile.puts "<p><p><h2 class=\"tabName\">#{document}</h2>"
    htmlFile.puts data[startOfHTML..-1]
  }
  htmlFile.puts "</body>\n</html>"
  htmlFile.close
end