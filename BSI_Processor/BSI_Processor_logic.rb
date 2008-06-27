require "BSI_Processor_GUI.rb"
require "BSI_Parser.rb"

class BSI_processor
  attr_reader :old_file_path, :new_file_path
  def init
    @option_excel.setCheck(false)
    @option_download.setCheck(false)
    
    @button_go.connect(Fox::SEL_COMMAND){
      start_processing
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
      end
    }
    
  @button_new_file_select.connect(Fox::SEL_COMMAND){
      if file_select_dialog.execute != 0
        @new_file_path = file_select_dialog.filename
        @new_file_name.text = @new_file_path.split("\\")[-1]
      end
    }
  end
  
  def start_processing
    @progressbar.total = 60000 
    @progressbar.progress = 0
    newData = BSI_excel_data.new(@new_file_path, self)
    oldData = BSI_excel_data.new(@old_file_path, self)
    newData.fixDates
    oldData.fixDates
    
    newItems = newData.getNewItems(oldData)
    if @option_download.checked?
      writeInfoPage(newItems.keys, "#{@output_filename_field.text}.html")
    end
    if @option_excel.checked?
      writeToExcel(prepareOutput(newItems),  "#{@output_filename_field.text}.xls")
    end
    @progressbar.progress = 60000
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