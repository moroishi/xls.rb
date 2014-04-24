#coding: cp932

require 'win32ole'

module Worksheet
  def [] y,x
    cell = self.Cells.Item(y,x)
    if cell.MergeCells
      cell.MergeArea.Item(1,1).Value
    else
      cell.Value
    end
  end

  def []= y,x,value
    cell = self.Cells.Item(y,x)
    if cell.MergeCells
      cell.MergeArea.Item(1,1).Value = value
    else
      cell.Value = value
    end
  end
end

class XlsApp
  def initialize()
    @xl = WIN32OLE.new('Excel.Application')
    @xl.Visible = true
    @xl.ScreenUpdating = false
  end

  def getAbsolutePath filename
    fso = WIN32OLE.new('Scripting.FileSystemObject')
    return fso.GetAbsolutePathName(filename)
  end

  def close
    @xl.ScreenUpdating = true
    @xl.Workbooks.each { |book| book.Close(false)  }  # force closing Workbooks
    @xl.Quit
  end

end

