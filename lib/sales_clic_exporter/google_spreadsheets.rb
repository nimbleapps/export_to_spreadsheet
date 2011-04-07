module SalesClicExporter::Document

  # Describes a Google Spreadsheets document
  class GoogleSpreadsheets < Base
    
    require 'google_spreadsheet' # Plug-in that handles Google Spreadsheets API communication
    require 'oauth'              # Plug-in that handles Google API authentication

    # Par défaut on commence à écrire sur la seconde colone.
    # A la différence de Excel, le premier index est 1 (au lieu de 0)
    Default_cell_index   = 2

    # Valeur qui va servir dans une route d'un controller à générer un document
    # Google Spreadsheets
    Extension            = 'google_spreadsheets'    

    def initialize(filename, args = {})
      # Cas username/password
      if args[:username] && args[:password]
        @username = args[:username]
        password  = args[:password]
        @google   = GoogleSpreadsheet.login(@username, password)
      # Token passé en paramètre
      elsif args[:access_token]
        token     = args[:access_token]
        @google   = GoogleSpreadsheet.login_with_oauth(token)
      elsif args.blank?
        @google   = GoogleSpreadsheet.saved_session
      else
        raise "Erreur lors de l'initialisation du module GoogleSpreadsheets"
      end

      @filename          = filename || "Export SalesClic"

      # Google Doc template, that contain styles
      @template_doc      = @google.spreadsheet_by_key(args[:google_doc_template_id]) if args[:google_doc_template_id]

      # If we can use a template we use it to create the file. Else we just create a new doc
      @doc               = @template_doc ? @template_doc.copy_doc(@filename) : @google.create_spreadsheet(@filename)

      # Ligne en cours d'écriture, afin de récupérer le dernier index écrit
      @current_row_index = 1 # Attention on démarre à la ligne 1
      
      # Réglage du nom de la première feuille du classeur
      sheet.title        = @filename
    end

    # Réglage du nom de la 1ère worksheet
    def worksheet_name=(name)
      sheet.title = name
    end

    # Nom de la 1ère worksheet
    def worksheet_name
      sheet.title
    end
    
    # Rajout d'une ligne vide
    def newline(size = nil)
      # Incrémentation du nombre de lignes dans le document
      @current_row_index += 1

      self
    end

    # Freezepane factice
    def freezepane
      self
    end

    # Sauvegarde du contenu
    def save
      sheet.save
    end

    # Suppression du document depuis Google Docs (vers la corbeille)
    def destroy
      @doc.delete
    end

    # Suppression du document depuis Google Docs (PERMANENTE)
    def permanent_destroy
      @doc.delete(true)
    end

    def key
      @doc.key
    end
    
    def url
      @doc.url
    end

    # On ne définit pas les méthodes title_x comme Spreadsheet ne supporte pas les titres.
    # Chaque méthode va en fait être utilisée comme une simple ligne.
    # Les titres 1 et 2 commencent à la première colone, tandis que le 3 est aligné avec le reste du texte
    def title_1(value)
      line(value, :index => 1)
    end

    def title_2(value)
      line(value, :index => 1)
    end

    def title_3(value)
      line(value, :index => Default_cell_index)
    end

    private

    # Ecriture d'une seule cellule
    def write_cell(value, cell_index = 1)
      sheet[@current_row_index, cell_index] = value
    end    
    
    # Accesseur en singleton pour @sheet.
    # S'il n'est pas défini, on l'assigne en vérifiant qu'il n'est pas nil.
    # Sinon on le renvoie après avoir vérifié AUSSI qu'il n'était pas nil.
    # Cette vérification permet de voir si on a pas une erreur liée au token
    # de connexion qui peut s'invalider
    def sheet
      if @sheet
        var_sheet = @sheet
      else
        var_sheet = @doc.worksheets[0]
      end
      
      if var_sheet.nil?
        raise "Erreur lors de la tentative d'écriture du document, le token n'est pas (plus) valide. Le problème peut venir des scopes passés à Google lors de la demande d'autorisation d'accès au compte de l'utilisateur (cf. http://code.google.com/apis/gdata/faq.html#AuthScopes)."
      else
        @sheet = var_sheet
      end
        
      return @sheet
    end    

    # Ecriture d'une ligne
    def write_line(values, cell_index, style_bold = nil, style_wrap = nil)
      # Cas de plusieurs valeurs
      if values.respond_to?('each') && values.size > 1
        values.each do |value|
          write_cell(value, cell_index) if value
          cell_index += 1
        end
      else
        # Gère les cas Ruby 1.8 et Ruby 1.9 (1.8, String répond à each, 1.9 non)
        value = values.respond_to?('each')   ?   values.first   :   values
        # Si on a qu'une seule valeur on l'écrit
        write_cell(value, cell_index) if value
      end

      return
    end    

  end

end