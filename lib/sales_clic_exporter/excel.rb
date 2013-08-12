# -*- encoding : utf-8 -*-
module SalesClicExporter::Document

  # Describes an Excel document
  class Excel < Base

    require 'rjb'

    # Ensures that the filename on the server is uniq
    attr_accessor :token

    # Par défaut on commence à écrire sur la colone 1 au lieu de 0
    # (sauf pour les titres)
    Default_cell_index         = 1

    # Taille des colones par défaut, en caractères (sauf la 1ère)
    # Dans la classe on se sert de @default_column_width qui peut prendre
    # cette valeur ou celle passée en option lors de l'instanciation
    Default_column_width       = 23
    
    # Taille de la première colone (qui sert à décaler le texte, pour
    # un effet de présentation
    Default_first_column_width = 5

    # Extension used to generate an Excel file
    Extension                  = 'xls'

    def initialize(client_filename, token = '')
      self.token                = token
      @client_filename          = client_filename
      @server_filename          = client_filename + token
    end

    # This is where the magic happens. It's called from #compute_export in export_to_spreadsheet.rb
    def finish_initialize(options = {})
      # Chargement de la MVJ
      # La tableau désigne les arguments passés en ligne de commande à la MVJ
      filedir = File.dirname(__FILE__) + '/..'
      memory  = ['-Xms256M', '-Xmx512M'] # Memory settings
      Rjb::load("#{filedir}/apache-poi/poi-3.7-20101029.jar", memory)

      # Import des packages Java
      @file_class               = Rjb::import('java.io.FileOutputStream')
      @workbook_class           = Rjb::import('org.apache.poi.hssf.usermodel.HSSFWorkbook')
      @cell_style_class         = Rjb::import('org.apache.poi.hssf.usermodel.HSSFCellStyle')
      @cell_range_address_class = Rjb::import('org.apache.poi.ss.util.CellRangeAddress')
      @region_util_class        = Rjb::import('org.apache.poi.hssf.util.HSSFRegionUtil')
      @font_class               = Rjb::import('org.apache.poi.hssf.usermodel.HSSFFont')

      # Création fichier et feuilles de calcul
      @book = @workbook_class.new
      # Reglage du nom de la première feuille
      @sheet = @book.createSheet(filename.capitalize)

      # Ligne en cours d'écriture, afin de récupérer le dernier index écrit
      @current_row_index = 0

      # Créé plusieurs style de formattage de manière statique.
      # Il faut faire attention à ne pas créer les style dans une boucle, leur nombre
      # étant limité. C'est pour cela qu'on les créé une fois, et qu'on ne fait ensuite
      # que les assigner à des cellules
      create_styles(options)

      # Application du formattage par défaut en fonction du Hash d'options
      # - réglage taille des colones
      # - réduction taille de la première colone
      default_formatting(options)
    end

    # Réglage du nom de la 1ère worksheet
    def worksheet_name=(name)
      @book.setSheetName(0, name)
    end

    # Nom de la 1ère worksheet
    def worksheet_name
      @book.getSheetName(0)
    end

    # Titre principal du document
    def title_1(text)
      title(text, 1)
    end

    # Titre secondaire
    def title_2(text)
      title(text, 2)
    end

    # Titre de niveau 3
    # Par défaut il commence sur la cellule d'index 1, au lieu de 0, car il n'est utilisé
    # que dans le "corps" de la feuille
    def title_3(text)
      title(text, 3)
    end

    # Rajout d'une ligne vide
    def newline(size = nil)
      if size
        row = @sheet.createRow(@current_row_index)
        row.setHeight(size * 20) # Configuration de la hauteur en 1/20 de caractère
      end

      increment_row_index

      self
    end

    # Créé un freezepane horizontal dans le document (fige la partie supérieure,
    # la partie inférieure est toujours srollable)
    # (fonction scinder/fixer sous Excel/Openoffice)
    def freezepane
      @sheet.createFreezePane(0, @current_row_index)
      self
    end

    # Nom du document sur le disque
    def filename
      "#{@server_filename}.#{Extension}"
    end

    # Nom 'formatté' du document, fournis à l'utilisateur
    def name
      "#{@client_filename}.#{Extension}"
    end    
    
    # Path du document
    def path
      root_path = defined?(Rails) ? "#{Rails.root}/" : "/"
      "#{root_path}tmp/#{filename}"
    end

    # Print du document
    def save
      # Ecriture des données dans un fichier du répertoire /tmp de Rails
      out = @file_class.new(self.path)
      @book.write(out)
      out.close
    end

    private

    # Ecriture d'une seule cellule
    def write_cell(value, row, cell_index = 0, style_name = nil)
      begin
        cell = row.createCell(cell_index)
        cell.setCellValue(value)
      rescue Exception => e
        raise %!Impossible d'écrire la valeur "#{value}" (#{value.class}) : #{e}!
      end

      begin
        if style_name
          # On recompose le nom de la variable d'instance contenant le style
          # à partir du nom passé en paramètre
          instance_variable_style = "@style_#{style_name}"
          # On récupère sa valeur
          style                   = instance_variable_get(instance_variable_style)
          # Assignation
          cell.setCellStyle(style) if style
        else
          cell.setCellStyle(instance_variable_get('@default_style'))
        end
      rescue
        raise %!Impossible de régler le style "#{style_name || default}" pour la valeur "#{value}" : #{e}!
      end
    end

    # Ecriture d'une ligne
    def write_line(values, cell_index, style)
      row = @sheet.createRow(@current_row_index)
      
      # Cas de plusieurs valeurs
      if values.respond_to?('each') && values.size > 1
        values.each do |value|
          write_cell(value, row, cell_index, style) if value
          cell_index += 1
        end
      else
        # Gère les cas Ruby 1.8 et Ruby 1.9 (1.8, String répond à each, 1.9 non)
        value = values.respond_to?('each')   ?   values.first   :   values
        # Si on a qu'une seule valeur on l'écrit
        write_cell(value, row, cell_index, style) if value
      end
      
      # Redimensionnage de la ligne
      if style && style.include?('wrap')
        # Taille par défaut des colones
        col_width   = @default_column_width # en caractères
        height      = @sheet.getDefaultRowHeightInPoints
        #debugger
        # On compte le nombre max de retours à la ligne dans les valeurs que l'on écrit
        lines = values.map do |v|
          nb_lines_per_value = []
          v.each_line do |elt|
                         size                = elt.size - 1 # On ne compte pas le \n
                         # Si le nombre de chars sur la ligne est plus petit que la largeur de la colone
                         # on compte 1, sinon on prend le rapport arrondi à l'entier supérieur
                         nb_lines_per_value << (size <= col_width ? 1 : ((size.to_f / col_width).ceil))
                       end
          # Additionne le nombre de retours comptés par case
          nb_lines_per_value.sum
        end.max # Et on prend le max pour avoir la valeur de retours la plus grande de la ligne
        
        row.setHeightInPoints(lines * height)
      end
      
      return
    end

    # Ajout d'une bordure en bas des cellules d'entête
    def add_bottom_border(headers)
      size_of_range      = headers.size - 1
      range_end_index    = size_of_range + Default_cell_index # Index de la fin de la région sur la ligne
      # On met la bordure à la ligne précédente
      row_index          = @current_row_index - 1
      cell_range_address = @cell_range_address_class.new(row_index, row_index, Default_cell_index, range_end_index)
      border             = @cell_style_class.BORDER_THIN
      @region_util_class.setBorderBottom(border, cell_range_address, @sheet, @book)
      return
    end

    # Réglages des styles des cellules
    def create_styles(options)
      # Police par defaut
      default_font   = @book.createFont
      default_font.setFontName   'Verdana'

      @default_style = @book.createCellStyle
      @default_style.setFont              default_font
      @default_style.setVerticalAlignment @cell_style_class.VERTICAL_TOP if options[:default_top_vertical_align]

      # Gras
      font           = @book.createFont
      font.setBoldweight         @font_class.BOLDWEIGHT_BOLD
      font.setFontName           'Verdana'

      @style_bold = @book.createCellStyle
      @style_bold.setFont        font
      @style_bold.setVerticalAlignment @cell_style_class.VERTICAL_TOP if options[:default_top_vertical_align]
      
      # Renvoie automatique à la ligne
      @style_wrap = @book.createCellStyle
      @style_wrap.setWrapText(true)
      @style_wrap.setVerticalAlignment @cell_style_class.VERTICAL_TOP if options[:default_top_vertical_align]
      
      # Couplage wrap + bold
      # Normalement on devrait pouvoir assigner plusieurs style
      # (un pour la cellule, l'autre pour la ligne.
      # Mais cela ne semble pas fonctionner pour le moment
      @style_bold_wrap = @book.createCellStyle
      @style_bold_wrap.setFont        font
      @style_bold_wrap.setWrapText(true)
      @style_bold_wrap.setVerticalAlignment @cell_style_class.VERTICAL_TOP if options[:default_top_vertical_align]
      
      # Titre 1
      font           = @book.createFont
      font.setFontName           'Verdana'
      font.setFontHeightInPoints 22

      @style_title_1 = @book.createCellStyle
      @style_title_1.setFont     font

      # Titre 2
      font           = @book.createFont
      font.setFontName            'Verdana'
      font.setFontHeightInPoints  18

      @style_title_2 = @book.createCellStyle
      @style_title_2.setFont      font

      # Titre 3
      font           = @book.createFont
      font.setFontName            'Verdana'
      font.setFontHeightInPoints  10
      font.setBoldweight          @font_class.BOLDWEIGHT_BOLD

      @style_title_3 = @book.createCellStyle
      @style_title_3.setFont      font

      return
    end

    # Ecriture d'un titre selon plusieurs paramètres
    def title(text, level)
      row        = @sheet.createRow(@current_row_index)

      # Décalage du titre 3 sur la cellule 1, comme il n'est utilisé que dans le texte
      cell_index =    level == 3   ?   1   :   0
      write_cell(text, row, cell_index, "title_#{level}".to_sym)

      # Incrémentation du nombre de lignes dans le document
      @current_row_index += 1

      self
    end

    # Formattage par défaut du workbook
    def default_formatting(options)
      @default_column_width       = options[:default_column_width] || Default_column_width
      
      # Taille des colones par défaut
      @sheet.setDefaultColumnWidth(@default_column_width)

      # Réglage de la taille de la première colone (en 256è de caractères)
      @sheet.setColumnWidth(0, Default_first_column_width * 256)
    end
    
  end

end
