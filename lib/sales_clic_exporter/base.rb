# Document describes a format to which data can be exported.
# A document can contain titles, lines or arrays.
module SalesClicExporter::Document

  # Gives the format that are supported by this plug-in. It can be used in routes
  # to check if the requested format can be generated.
  # @example
  #   SalesClicExporter::Document.allowed_extensions #=> ['xls', 'google_spreadsheet']
  def self.allowed_extensions
    subclasses = self::Base.subclasses
    subclasses.map do |str_klass|
      klass = str_klass.constantize
      klass.const_defined?(:Extension)   ?   klass::Extension   :   nil
    end.compact
  end

  # All formats will inherit from this class
  class Base
    
    # Ecriture d'un tableau vertical avec une colone d'entetes (verticalement)
    # et une autre colone de valeurs
    # Prend en 2è argument un tableau, ou une suite d'arguments
    def v_table(array_of_values, *headers)
      raise "array_of_values should be an array" unless array_of_values.kind_of?(Array)

      # Si on passe un tableau, headers prend comme valeur [valeurs du tableau], ie. il
      # encapsule les valeurs dans un tableau. Dans ce cas, on pend le premier élément
      headers       = headers.first if headers.is_a?(Array) && headers.size == 1 && headers.first.is_a?(Array)
      
      # On forme un zip des deux tableau, cf. la doc de Ruby (Array#zip).
      # La méthode agit comme une transposition (les entetes et les valeurs étant deux vecteurs)
      # sauf que les valeurs absentes DU SECOND TABLEAU (celui en paramètre de #zip)
      # sont remplacées par nil.
      
      # C'est pour cela qu'on va d'abord chercher quel est le tableau le plus petit
      # afin d'avoir l'ordre permettant que les 'trous' soient comblés (il faut que
      # ce tableau plus petit soit en paramètre).
      # Sinon le résultat est tronqué à la première valeur nil trouvée dans le premier tableau
      # Par la suite, on va dans ce cas (si l'ordre change) inverser les résultats
      # afin de ne pas avoir [VALEUR, COLONE] au lieu de [COLONE, VALEUR]
      ary_a         = headers.size < array_of_values.size   ?   array_of_values   :   headers
      ary_b         = headers.size < array_of_values.size   ?   headers           :   array_of_values
      zipped_values = *ary_a.zip(ary_b)

      #Handle the case where array_of_values and headers are Arrays of one unique element
      #In this case, zip does not return an array of array like this [[key,value],[key,value]]
      #But just a simple array [key,value]. In this case, we just write a line
      unless zipped_values.first.is_a? Array
        line(zipped_values).newline
        return self
      end

      zipped_values.each do |array_of_two_values|
        # C'est ici qu'on regarde si l'ordre a été changé
        # auquel cas on inverse les valeurs du tableau
        if headers.size < array_of_values.size
          ary = array_of_two_values.reverse
        else
          ary = array_of_two_values
        end
        line(ary)
      end

      self
    end

    # Méthode de complésance pour écrire dans le document en utilisant un bloc
    def write
      yield(self) if block_given?
      self
    end

    # Ecriture d'un tableau horizontal avec une matrice de valeurs, et les entetes
    # des colones (horizontalement)
    # Le second argument comprend les headers, et les options. Cette syntaxe permet une grande liberté d'écriture
    # car on peut spécifier une liste d'headers comme liste d'arguments, puis en dernier argument un hash d'options
    def h_table(values, *headers_with_options)
      raise ArgumentError, "values should be an array" unless values.kind_of?(Array)

      # Extraction du dernier argument
      options = (headers_with_options.last.kind_of?(Hash) && headers_with_options.respond_to?('each'))   ?   headers_with_options.pop   :   {}
      
      # Pour plus de lisibilité on renomme la variable à utiliser dans le reste de la méthode
      # Le * permet d'appliatir le tableau d'arguments, qui est contenu dans un tableau sinon (en fait, à ce moment là
      # headers_with_options (sans le *) vaut [['a', 'b', 'c']] par exemple
      headers = *headers_with_options

      # Dans la méthode #line on utilise le même mécanisme pour contenir les options dans la liste d'arguments.
      # C'est pour cette raison qu'on réintègre les options dans le header, et qu'on passe le tout en argument
      line(headers << options.merge({:bold => true, :wrap_text => true}))

      # Ajout d'une bordure sous les headers
      # respond_to? ne prend pas en compte les méthodes privée, on utilise à la place
      # private_methods qui renvoie la liste des méthodes privées
      if options[:border_bottom] && self.private_methods.include?('add_bottom_border')
        add_bottom_border(headers)
      end

      # Ligne de démarquation
      newline(6)

      values.each do |array_of_values|
        line(*array_of_values << options)
      end

      self # Pour pouvoir chaîner les méthodes
    end

    # Ecriture d'une ligne
    def line(*values)
      raise ArgumentError, "values ne peut être vide" if values.empty?

      # Gestion du cas où on passe un tableau plutôt que plusieurs arguments
      # Dans ce cas values = [['a', 'b', 'c']] par exemple, ie. le tableau
      # passé en paramètre se retrouve dans un tableau, dont on ne prend que le
      # premier élément
      if values.kind_of?(Array) && values.size == 1 && values.first.kind_of?(Array)
        values = values.first
      end

      # Options de formattage
      # On prend la dernière valeur du paramètre, si c'est bien un hash
      # On procède de cette manière pour pouvoir toujours utiliser la liste
      # d'arguments, mais sans spécifier d'options vide, qui serait obligatoire à l'appel
      # le cas échéant
      options = values.pop if values.last.kind_of?(Hash) && values.respond_to?('each')

      # Formattage
      style = if options && options[:bold] && options[:wrap_text]
                  'bold_wrap'
              elsif options && options[:wrap_text]
                  'wrap'
              elsif options && options[:bold]
                  'bold'
              else
                nil
              end                

      # Index de cellule spécifié ? (pour démarer la ligne sur la
      # celulle d'index 1 au lieu de 0)
      cell_index = options && options[:index]  ?   options[:index]   :   self.class::Default_cell_index

      write_line(values, cell_index, style)

      increment_row_index

      self # Chainage des méthodes d'écriture
    end    

    private
    
    # Incrémente le nombre de lignes du document
    def increment_row_index
      # Incrémentation du nombre de lignes dans le document
      @current_row_index += 1
      return
    end
    
  end

end