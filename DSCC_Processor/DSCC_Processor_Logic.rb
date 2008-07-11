require 'DSCC_Processor_GUI.rb'
require 'DSCC_Processor.rb'

class DSCC_Email_Processor
  def init
    @buttonGo.connect(Fox::SEL_COMMAND){
      dataArray = @textField.text.split("\n")
      @DSCC_data = EmailInformation.new(dataArray, self)
      @processingThread = Thread.new{
        @textField.setText("Preparing to download documents...\n")
        @DSCC_data.downloadDocuments("TestDownload")
      }
    }
    
    @buttonStop.connect(Fox::SEL_COMMAND){
      #Stuff to do when stop is pressed
    }
  end
  
  def processData
    
  end
end

#unit test
if __FILE__==$0
	require 'libGUIb16'
	app=FX::App.new
	w=DSCC_Email_Processor.new app
	w.topwin.show(0)
	app.create
  app.run
end