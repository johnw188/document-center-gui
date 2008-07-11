require 'net/http'
require 'ftools'

class EmailInformation
  
  attr_reader :emailFileName, :emailHash

  DOCUMENT, DATE, DESCRIPTION, URL, FILESIZE, MOREINFO = 0, 1, 2, 3, 4, 5

  def initialize(dataArray, gui = nil)
    puts "Initializing"
    @processorGUI = gui
    @fileArray = dataArray
    @brokenDownData = breakDownData(@fileArray)
    @emailHash = dataToHash(@brokenDownData)
  end
  
  def breakDownData(array)
    data = array.map {|line| line.strip}
    until data[0] =~ /Document:/
      data.shift
    end
    until data[-1] =~ /You can browse an index of all.*/
      data.pop
    end
    data.to_s.split(/(Document:|--  Dated|Description:|URL:|File size:|More info:)/)
  end

  def dataToHash(dataArray)
    defaultArray = ["NO DATA", "NO DATA", "NO DATA", "NO DATA", "NO DATA", "NO DATA"]
    dataHash = Hash.new
    dataArray.each_with_index do |item, index|
      if item =~ /Document:/
        documentName = dataArray[index + 1].strip
        dataHash[documentName] = defaultArray.dup
        dataHash[documentName][DOCUMENT] = documentName
        for i in 2..11
          break if dataArray[index + i] =~ /Document:/
          case dataArray[index + i]
          when /--  Dated/
            dataHash[documentName][DATE] = dataArray[index + i + 1].strip
          when /Description:/
            dataHash[documentName][DESCRIPTION] = dataArray[index + i + 1].strip
          when /URL:/
            dataHash[documentName][URL] = dataArray[index + i + 1].strip
          when /File size:/
            dataHash[documentName][FILESIZE] = dataArray[index + i + 1].strip
          when /More info:/
            dataHash[documentName][MOREINFO] = dataArray[index + i + 1].strip
          end
        end
      end
    end
    dataHash
  end
    
    
  def getURLs
    @emailHash.collect {|key,value| value[URL]}
  end

  def getDocNames
    @emailHash.collect {|key,value| value[DOCUMENT]}
  end
  
  def getDocInfo(documentName)
    @emailHash[documentName]
  end

  def downloadDocuments(path = -1)
    workingDirectory = Dir.getwd
    unless path == -1
      if !Dir[path].any? then Dir.mkdir(path) end
      Dir.chdir(path)
    end
    
    numberOfFiles = @emailHash.keys.size
    i = 0

    h = Net::HTTP.start("www.dscc.dla.mil")
    @emailHash.keys.each do |documentName|
      @processorGUI.textField.appendText("#{documentName}...")
      path = @emailHash[documentName][URL][/\/Downloads\/.*/]
      resp = h.get(path)
      File.open(documentName.gsub(/ /, '_').gsub(/[\/\\]/, '+').gsub(/_\(/, '').gsub(/\)\./, '.').gsub(/Rev/, '_Rev').gsub(/Initial/, '_Initial').gsub(')', '') + path[/\.[a-z]+/], "wb") {|file| file.write(resp.body)}
      print documentName, " was downloaded successfully\n"
      i += 1
      if i == numberOfFiles
        @processorGUI.textField.appendText(" Downloaded!\nFinished processing, you may exit at any time")
      else
        @processorGUI.textField.appendText(" Downloaded! (#{numberOfFiles - i} remaining)\n")
      end
    end
    
    unless path == -1
      Dir.chdir(workingDirectory)
    end
  end
  
end

class EmailUILoop
  def startUI
    files = Dir["*"]

    puts "This program processes the email data we receive from DSCC.\n\nTo use it, first open Mozilla Thunderbird, select the email to be processed,\nand choose \"save as\". Make sure that the email file is in the same folder\nas this script. You will be prompted for a folder name to save the\nfiles to, and then the script will download all of the PDF's linked to\nin that email. After that, you simply need to select all the files, open them,\nmaximize the window, and run the JournalMacro \"Print first page\" script\nuntil the first page of each document has been printed. This data goes to\ndata entry.\n\n"
    puts "Enter a folder name to download the files into"
    folderName = gets.chomp

    again = true
    while again
      label = "a"
      puts "Files in this folder: "

      files.each do |filename|
	puts label + ") " + filename
	label.next!
      end

      choices = ("a"...label).to_a

      puts "Which file do you want to process (a - " + choices[-1] + ")?"
      userChoice = gets.strip

      puts "Processing information in the file " + files[choices.rindex(userChoice)] 
      puts "This may take a moment...\n"

      EmailInformation.new(files[choices.rindex(userChoice)]).downloadDocuments(folderName)
      print "\n\nProcess another file? [y/n] "
      again = false if gets.chomp.downcase == "n"
    end
  end
end

class MoveAndRename
  def initialize(path)
    @path = path
    Dir.chdir(path)
    @files = Dir["*"]
  end
end