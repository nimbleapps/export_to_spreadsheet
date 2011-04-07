require 'test_helper'

class GoogleSpreadsheetsTest < ActiveSupport::TestCase
  
  Sample_string = 'Test chaîne'
  
  if ! File.file?(ENV["HOME"] + '/.ruby_google_spreadsheet.token')
    puts "Pour la connexion à Google vous devez indiquer votre email de connexion et votre mot de passe."
    puts "Les tokens générés sont stockés dans le fichier ~/.ruby_google_spreadsheet.token que vous pouvez supprimer."
  end
  
  def setup
    @obj    = create_obj    
    @google = SalesClicExporter::Document::GoogleSpreadsheets.new @obj.name
  end

  def teardown
    # Sauvegarde du document
    @google.save
    # Suppression des feuilles crées
    @google.permanent_destroy
  end
  
  ###################################
  # Tests sur les arguments
  ###################################

  def test_line_with_and_without_splat_arguments
    assert_method_returns_self('line', 'colone 1', 'colone 2', 'colone 3', 'colone 4')
    assert_method_returns_self('line', ['colone 1', 'colone 2', 'colone 3', 'colone 4'])
    assert_method_returns_self('line', 'colone 1')
  end

  def test_line_without_arguments_throws_exception
    assert_raise ArgumentError do
      @google.line
    end
  end

  def test_h_table_without_values_as_an_array_throws_exception
    assert_raise ArgumentError do
      @google.h_table('test', 'test')
    end
  end

  def test_h_table_with_and_without_options_and_splat_arguments
    assert_method_returns_self('h_table', [%w{a b c}, %w{d e f}, %w{g h}], ['toto', 'tata', 'tutu'])
    assert_method_returns_self('h_table', [%w{a b c}, %w{d e f}, %w{g h}], ['toto', 'tata', 'tutu'], {:border_bottom => true})
    assert_method_returns_self('h_table', [%w{a b c}, %w{d e f}, %w{g h}], 'toto', 'tata', 'tutu')
    assert_method_returns_self('h_table', [%w{a b c}, %w{d e f}, %w{g h}], 'toto', 'tata', 'tutu', {:border_bottom => true})
  end

  def test_v_table_with_and_without_splat_arguments
    assert_method_returns_self('v_table', %w{toto tata tutu}, 'colone 1', 'colone 2', 'colone 3', 'colone 4')
    assert_method_returns_self('v_table', %w{toto tata tutu}, ['colone 1', 'colone 2', 'colone 3', 'colone 4'])
  end

  def test_newline_can_be_called_without_arguments
    assert_nothing_raised do
      @google.newline
    end
  end

  def test_write_method_without_block_is_ignored
    assert_equal @google.write, @google
  end


  ######################################################################
  # On teste les retours pour vérifier que les méthodes sont bien
  # chainables
  ######################################################################

  def test_line_returns_self
    assert_method_returns_self('line', 'colone 1', 'colone 2', 'colone 3', 'colone 4')
  end

  def test_h_table_returns_self
    assert_method_returns_self('h_table', [%w{a b c}, %w{d e f}, %w{g h}], ['toto', 'tata', 'tutu'], {:border_bottom => true})
  end

  def test_v_table_returns_self
    assert_method_returns_self('v_table', %w{toto tata tutu}, 'colone 1', 'colone 2', 'colone 3', 'colone 4')
  end

  def test_newline_returns_self
    assert_method_returns_self('newline')
  end

  def test_freezepane_returns_self
    assert_method_returns_self('freezepane')
  end

  def test_title_1_returns_self
    assert_method_returns_self('title_1', Sample_string)
  end

  def test_title_2_returns_self
    assert_method_returns_self('title_2', Sample_string)
  end

  def test_title_3_returns_self
    assert_method_returns_self('title_3', Sample_string)
  end

  ######################################################################
  # On teste sur toutes les méthodes rajoutent bien une ligne
  ######################################################################

  def test_h_table_increments_current_row_index
    headers     = ['toto', 'tata', 'tutu']
    lines_added = headers.map(&:size).max + 1 # + 1 pour la ligne vide de séparation
    assert_method_increments_current_row_index(lines_added, 'h_table', [%w{a b c}, %w{d e f}, %w{g h}], ['toto', 'tata', 'tutu'], {:border_bottom => true})
  end

  def test_v_table_increments_current_row_index
    headers     = %w{toto tata tutu}
    values      = ['colone 1', 'colone 2', 'colone 3', 'colone 4']
    lines_added = [values.size, headers.size].max
    assert_method_increments_current_row_index(lines_added, 'v_table', headers, *values)
  end

  def test_line_increments_current_row_index
    assert_method_increments_current_row_index(1, 'line', 'colone 1', 'colone 2', 'colone 3', 'colone 4')
  end

  def test_newline_increments_current_row_index
    assert_method_increments_current_row_index(1, 'newline')
  end

  # Freezepane ne DOIT pas incrémenter le nombre de lignes
  def test_freezepane_does_not_increment_current_row_index
    row_index      = @google.current_row_index
    index_expected = row_index # Le nombre de lignes ne doit PAS changer

    @google.freezepane

    assert_equal index_expected, @google.current_row_index
  end
  
  #########################################################
  # Test de lecture des données sur le document spreadsheet
  #########################################################

  def test_title_1_writes_correct_values
    text   = 'Titre looooooooooong'
    # Valeur attendue
    values = [ expected_title(text) ]
    assert_write_correct_value_with_function(values, 'title_1', text)
  end

  def test_title_2_writes_correct_values
    text   = 'Titre looooooooooong'
    values = [ expected_title(text) ]
    assert_write_correct_value_with_function(values, 'title_2', text)
  end

  def test_title_3_writes_correct_values
    text   = 'Titre looooooooooong'
    values = [ expected_subtitle(text) ]
    assert_write_correct_value_with_function(values, 'title_3', text)
  end

  def test_line_writes_correct_values
    # Valeurs attendues
    values = [ expected_text('Valeurs', 'sur', 'une', 'ligne') ]
    assert_write_correct_value_with_function(values, 'line', *['Valeurs', 'sur', 'une', 'ligne'])
  end

  def test_h_table_writes_correct_values
    # Valeurs attendues
    values =  []
    values += expected_table([
                ['colone 1', 'colone 2', 'colone 3'],
                [nil],
                ['a', 'b', 'c'],
                [nil, nil, 'd'],
                ['e'],
                [nil, 'f']
              ])
    # Ce que j'écris dans le fichier
    headers = ['colone 1', 'colone 2', 'colone 3']
    write   = [ ['a', 'b', 'c'], [nil, nil, 'd'], ['e'], [nil, 'f', nil] ]
    assert_write_correct_value_with_function(values, 'h_table', write, headers, {:border_bottom => true})
  end

  # Il n'y a pas de test où h_table est appelé avec :border_bottom => false
  # car la ligne de démarquation est ajoutée dans tous les cas

  def test_v_table_writes_correct_values
    values = []
    values += expected_table([
                ['colone 1', 'valeur 1'],
                ['colone 2', 'valeur 2'],
                ['colone 3'            ],
                [       nil, 'valeur 3'],
                ['colone 4', 'valeur 4']])
    headers = ['colone 1', 'colone 2', 'colone 3', nil, 'colone 4']
    write   = ['valeur 1', 'valeur 2', nil, 'valeur 3', 'valeur 4']
    assert_write_correct_value_with_function(values, 'v_table', write, headers)
  end

  def test_complete_write_document
    @google.title_1('Titre 1')
    @google.title_2('Titre 2')
    @google.title_3('Titre 3')
    @google.line("Ligne de test")
    @google.h_table([%w{a b c}, %w{d e f}, %w{g h}], ['toto', 'tata', 'tutu'], {:border_bottom => true})
    @google.v_table(%w{toto tata tutu}, ['colone 1', 'colone 2', 'colone 3', 'colone 4'])
    @google.newline
    @google.line("Autre ligne de test")
    @google.save   
    
    expected_values  = []
    expected_values << expected_title('Titre 1')
    expected_values << expected_title('Titre 2')
    expected_values << expected_subtitle("Titre 3")
    expected_values << expected_text("Ligne de test")
    expected_values << expected_text('toto', 'tata', 'tutu')
    expected_values << [nil]
    expected_values += expected_table([['a', 'b', 'c'], ['d', 'e', 'f'], ['g', 'h']])
    expected_values += expected_table([['colone 1', 'toto'], ['colone 2', 'tata'], ['colone 3', 'tutu'], ['colone 4']])
    expected_values << [nil]
    expected_values << expected_text("Autre ligne de test")
    
    assert_equal expected_values, array_from_google_spreadsheet
  end  
  
  private

  def assert_method_returns_self(method, *args)
    return_class = @google.send(method, *args).class

    assert_equal @google.class, return_class
  end

  def assert_method_increments_current_row_index(number_of_lines_added, method, *args)
    row_index      = @google.current_row_index
    index_expected = row_index + number_of_lines_added

    @google.send(method, *args)

    assert_equal index_expected, @google.current_row_index
  end
  
  def assert_write_correct_value_with_function(expected_values, method, *args)
    @google.send(method, *args)    
    
    @google.save
    
    assert_equal expected_values, array_from_google_spreadsheet
  end
  
  def array_from_google_spreadsheet
    # On travaille sur la première spreadsheet enregistrée, et sur sa
    # première worksheet
    sheet = @google.google_doc.spreadsheets.first.worksheets.first
    
    values = []
    for row_index in 1..sheet.max_rows    
      row_values = []
      for col_index in 1..sheet.max_cols
        cell        = sheet[row_index, col_index]
        row_values << (cell.blank?   ?   nil   :   cell)        
      end
      # Suppression des éléments nil après le dernier élément non nil
      values    << remove_nil_elements_from_end_of_array(row_values)
    end   
    
    # On supprime les [nil] en double en fin de tableau
    remove_nil_arrays_from_array(values)
  end
  
  # Soit un tableau qui comprend des valeurs nil et non-nil. La fonction retire
  # les éléments nil "en trop" depuis la fin du tableau (en s'arrêtant au premier élément
  # non-nil depuis la fin du tableau)
  def remove_nil_elements_from_end_of_array(ary)
    return [nil] if ary.compact.size == 0
    
    reverse_ary                = ary.reverse
    last_non_nil_element       = reverse_ary.compact.first
    last_non_nil_element_index = reverse_ary.index(last_non_nil_element)
    non_nil_elements           = reverse_ary.slice(last_non_nil_element_index..reverse_ary.size)
    non_nil_elements.reverse
  end

  # Soit un tableau contenant des tableaux contenant des valeurs (nil et non-nil) ou seulement nil.
  # La fonction retire les éléments [nil] "en trop" depuis la fin du tableau (en s'arrêtant au premier
  # élément diférent de [nil] depuis la fin du tableau)
  def remove_nil_arrays_from_array(ary)
    return nil if ary.compact.size == 0
    
    reverse_ary                = ary.reverse
    # On cherche le dernier élément qui n'est pas égal à [nil]
    # donc on prend PAS ceux qui ne contiennent qu'un élément et dont cet élément serait nil
    # et on garde le dernier élément
    last_non_nil_element       = ary.select{ |v| v.size != 1 || ! v.first.nil? }.last
    last_non_nil_element_index = reverse_ary.index(last_non_nil_element)
    non_nil_elements           = reverse_ary.slice(last_non_nil_element_index..reverse_ary.size)
    non_nil_elements.reverse
  end

  # When a lines contains a title, it should be written on the first cell
  def expected_title(title)
    title.to_a
  end


  # When a lines contains a subtitle title_3), it should be written on the default cell
  def expected_subtitle(title)
    # Returns an Array with nil cells merged with an Array containing the title
    array_with_nil_cells + title.to_a
  end

  # When a lines contains a text, it should be written from the second cell
  def expected_text(*text)
    array_with_nil_cells + text
  end

  # When many lines contain a table, it should be written from the second column
  def expected_table(ary)
    ary.map do |line|
      # An empty line should just have nil once
      line.size == 1 && line.first.nil? ? line : (array_with_nil_cells + line)
    end
  end

  def array_with_nil_cells
    Array.new(cell_index)
  end

  def cell_index
    # The cell index in spreadsheets starts from 1, and Ruby starts from 0
    SalesClicExporter::Document::GoogleSpreadsheets::Default_cell_index - 1
  end
  
end