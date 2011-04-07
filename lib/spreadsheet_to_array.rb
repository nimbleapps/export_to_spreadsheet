# This module intent to convert a spreadsheet into a matrix (Array of Array).
# Currently only Excel is supported, Google Spreadheets data can be read directly
# but a wrapper could be added in this module.
#
# @example
#   filepath = 'test_file.xls'
#   bytes    = IO::read(filepath)
#   excel    = SpreadsheetToArray::FromExcel.new(bytes)
#   puts excel.values
module SpreadsheetToArray
  
  class FromExcel
    
    attr_reader :values
    
    def initialize(byte_array)
    
      require 'rjb'
      
      # JVM loading
      filedir = File.dirname(__FILE__) + '/.'
      # Memory settings
      memory = ['-Xms256M', '-Xmx512M']
      Rjb::load("#{filedir}/apache-poi/poi-3.7-20101029.jar", memory)
      
      # Import des packages Java
      begin
        @file_class               = Rjb::import('java.io.FileInputStream')
        @workbook_class           = Rjb::import('org.apache.poi.hssf.usermodel.HSSFWorkbook')
        @cell_class               = Rjb::import('org.apache.poi.hssf.usermodel.HSSFCell')
        @date_util_class          = Rjb::import('org.apache.poi.ss.usermodel.DateUtil')
        @byte_array_input_stream  = Rjb::import('java.io.ByteArrayInputStream')
      rescue
        raise "Impossible to load Java packages. Maybe the path of Apache POI is not correct."
      end      
      
      begin
        java_byte_array         = @byte_array_input_stream.new_with_sig('[B', byte_array)
        @book                   = @workbook_class.new(java_byte_array)
      rescue Exception => e
        raise "The document is not a valid Excel file: #{e}"
      end
      
      # Reglage du nom de la première feuille
      @sheet                    = @book.getSheetAt(0)
      
      # Valeurs récupérées depuis le tableur
      @values                   = []
      parse_sheet    
    end
    
    private
    
    # Parcours la feuille et enregistre les valeurs trouvées dans la structure adécouate
    def parse_sheet   
      
      # On commence à la 1è ligne
      @current_row_index = 0       
      
      # Parcours de l'intervalle de la première ligne à la dernier du document
      for row_index in @current_row_index..@sheet.getLastRowNum.to_i
        row  = @sheet.getRow(row_index)
        parse_row(row)        
      end
      
    end # fin fonction
    
    # Dans une ligne (à partir de la 3è), chaque valeur est enregistrée dans un hash dont la clé
    # correspont à la valeur de la 1ère colone {:objectif => 9000}, lui-même associé à la clé
    # correspondant à la période {:period_x => {:objectif => 9000, :realise => 6000}}
    def parse_row(row)
      # Valeurs de la ligne
      row_values = []      
      
      # On passe à la ligne d'après s'il n'y a aucune cellule dans la ligne
      if row.nil? || row.getPhysicalNumberOfCells == 0        
        @current_row_index += 1
        @values << [nil]
        
        return
      end
      
      # Parcours de l'intervalle qui couvre l'intervalle de 1 (la cellule 0 correspond au label)
      # au nombre de cellules max (borne ouverte) de la ligne
      for cell_index in 0...row.getLastCellNum.to_i
        cell        = row.getCell(cell_index)
        row_values << parse_cell(cell)
      end
      
      @values << row_values
      
      @current_row_index += 1
        
      return      
    end
    
    # Lecture d'une cellule
    def parse_cell(cell)
      # Si la cellule est vide on retourne nil
      return if cell.nil?
      
      case cell.getCellType
      when @cell_class.CELL_TYPE_STRING
        cell.getRichStringCellValue.getString
      when @cell_class.CELL_TYPE_NUMERIC
        if @date_util_class.isCellDateFormatted(cell)
          cell.getDateCellValue.toString
        else
          cell.getNumericCellValue
        end
      else
        nil
      end      
    end     
  
  end
  
end