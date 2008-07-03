require "BSI_Processor_GUI.rb"
require "BSI_Parser.rb"

class BSI_processor
  attr_reader :old_file_path, :new_file_path
  def init
    @option_excel.setCheck(false)
    @option_download.setCheck(false)
    
    @oldFile = ""
    @newFile = ""
    @userDefinedFilename = false
    
    @button_go.connect(Fox::SEL_COMMAND){
      if (@option_excel.unchecked? && @option_download.unchecked?)
        Fox::FXMessageBox.warning(@topwin, Fox::MBOX_OK, "No Output Selected", "You must select at least one data output option")
      elsif (@output_filename_field.text == "")
        Fox::FXMessageBox.warning(@topwin, Fox::MBOX_OK, "No Filename", "You must enter a filename for the output files")
      else
        @processingThread = Thread.new{
          @output_text.appendText("\nLoading requested files, please be patient...")
          start_processing
        }
      end
    }
    
    @button_stop.connect(Fox::SEL_COMMAND){
      @processingThread.kill
      @progressbar.progress = 0
      @progressbar.showNumber
    }
    
    file_select_dialog = Fox::FXFileDialog.new(@topwin, "Select a BSI data extract")
    
    file_select_dialog.patternList = [
      "All Files(*)",
      "Excel Files (*.xls)"
    ]
    
    file_select_dialog.selectMode = Fox::SELECTFILE_EXISTING
    
    @button_old_file_select.connect(Fox::SEL_COMMAND){
      if file_select_dialog.execute != 0
        @old_file_path = file_select_dialog.filename
        @old_file_name.text = @old_file_path.split("\\")[-1]
        @oldFile = @old_file_name.text.split(".")[0].split(" ")[-1]
        if @userDefinedFilename == false
          @output_filename_field.text = "BSI_diff_" + @oldFile + "_" + @newFile
        end
      end
    }
    
  @button_new_file_select.connect(Fox::SEL_COMMAND){
      if file_select_dialog.execute != 0
        @new_file_path = file_select_dialog.filename
        @new_file_name.text = @new_file_path.split("\\")[-1]
        @newFile = @new_file_name.text.split(".")[0].split(" ")[-1]
        if @userDefinedFilename == false
          @output_filename_field.text = "BSI_diff_" + @oldFile + "_" + @newFile
        end
      end
    }
    
    @output_filename_field.connect(Fox::SEL_CHANGED){
      @userDefinedFilename = true
    }
    
    @output_filename_field.connect(Fox::SEL_COMMAND){|sender, selector, data|
      if data =~ /[,|.|;|'|"|`|~|=|:|\|]/
        Fox::FXMessageBox.warning(@topwin, Fox::MBOX_OK, "Filename Error", "Filename contains invalid characters, please fix it")
        @output_filename_field.selectAll()
      end
    }
  end
  
  def start_processing
    @progressbar.total = 60000 
    @progressbar.progress = 0
    @progressbar.showNumber
    newData = BSI_excel_data.new(@new_file_path, self)
    oldData = BSI_excel_data.new(@old_file_path, self)
    newData.fixDates
    oldData.fixDates
  
    newItems = newData.getNewItems(oldData)
    if @option_excel.checked?
      @output_text.appendText("\nWriting Excel spreadsheet to #{@output_filename_field.text}.xls...")
      writeToExcel(prepareOutput(newItems),  "#{@output_filename_field.text}.xls")
      @output_text.appendText(" Success!")
    end
    if @option_download.checked?
      @output_text.appendText("\nWriting HTML report to #{@output_filename_field.text}.html...")
      writeInfoPage(newItems.keys, "#{@output_filename_field.text}.html")
      @output_text.appendText(" Success!")
    end
    @progressbar.progress = 59999
    @progressbar.increment(1)
    @progressbar.showNumber
    @output_text.appendText("\nFinished processing data")
  end
end

#unit test
if __FILE__==$0
	require 'libGUIb16'
	app=FX::App.new
	w=BSI_processor.new app
	w.topwin.show(0)
	app.create
  app.run

end