require "BSI_Processor_GUI.rb"

class MainWindow
  def init
    @button_go.connect(Fox::SEL_COMMAND){
      @text2.text = @text2.text + "\nNewLine"
    }
  end
end

#unit test
if __FILE__==$0
	require 'libGUIb16'
	app=FX::App.new
	w=MainWindow.new app
	w.topwin.show(0)
	app.create
	app.run

end