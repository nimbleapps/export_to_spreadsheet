# -*- encoding : utf-8 -*-
require 'test_helper'

class ExcelTest <  ActiveSupport::TestCase
  Sample_string = 'Test chaîne'
  
  def setup
    @obj   = create_obj
    #The export initialization process is done in two steps because of rjb which takes a lot of memory :
    #First we set some variables with the private method prepare_export. We use send to be able to clal this method from outside only in the tests
    @excel = @obj.send(:prepare_export, :excel)
    #Then we fork the process when possible and finally load rjb
    #with the method finish_initialize
    @excel.finish_initialize
  end
  
  def teardown
    if File.file?(@excel.path)
      assert File.unlink(@excel.path)
    end
  end

  ###################################
  # Tests sur les arguments
  ###################################

  def test_write_method_without_block_is_ignored
    assert_equal @excel.write, @excel
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

  def test_line_with_and_without_splat_arguments
    assert_method_returns_self('line', 'colone 1', 'colone 2', 'colone 3', 'colone 4')
    assert_method_returns_self('line', ['colone 1', 'colone 2', 'colone 3', 'colone 4'])
    assert_method_returns_self('line', 'colone 1')
  end

  def test_line_without_arguments_throws_exception
    assert_raise ArgumentError do
      @excel.line
    end
  end

  def test_h_table_without_values_as_an_array_throws_exception
    assert_raise ArgumentError do
      @excel.h_table('test', 'test')
    end
  end

  def test_newline_can_be_called_without_arguments
    assert_nothing_raised do
      @excel.newline
    end
  end  

  ######################################################################
  # On teste les retours pour vérifier que les méthodes sont bien
  # chainables
  ######################################################################

  def test_title_1_returns_self
    assert_method_returns_self('title_1', Sample_string)
  end

  def test_title_2_returns_self
    assert_method_returns_self('title_2', Sample_string)
  end

  def test_title_3_returns_self
    assert_method_returns_self('title_3', Sample_string)
  end

  def test_h_table_returns_self
    assert_method_returns_self('h_table', [%w{a b c}, %w{d e f}, %w{g h}], ['toto', 'tata', 'tutu'], {:border_bottom => true})
  end

  def test_v_table_returns_self
    assert_method_returns_self('v_table', %w{toto tata tutu}, 'colone 1', 'colone 2', 'colone 3', 'colone 4')
  end

  def test_line_returns_self
    assert_method_returns_self('line', 'colone 1', 'colone 2', 'colone 3', 'colone 4')
  end

  def test_newline_returns_self
    assert_method_returns_self('newline')
  end

  def test_freezepane_returns_self
    assert_method_returns_self('freezepane')
  end

  def test_write_returns_self
    assert_method_returns_self('write')
  end

  ######################################################################
  # On teste sur toutes les méthodes rajoutent bien une ligne
  ######################################################################
  
  def test_title_1_increments_current_row_index
    assert_method_increments_current_row_index(1, 'title_1', Sample_string)
  end

  def test_title_2_increments_current_row_index
    assert_method_increments_current_row_index(1, 'title_2', Sample_string)
  end

  def test_title_3_increments_current_row_index
    assert_method_increments_current_row_index(1, 'title_3', Sample_string)
  end

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

  def test_freezepane_does_not_increment_current_row_index
    row_index      = @excel.current_row_index
    index_expected = row_index # Pas de changement

    @excel.freezepane

    assert_equal index_expected, @excel.current_row_index    
  end
  
  ##########################################
  # Tests sur la sauvegarde
  ##########################################

  def test_save
    @excel.line("ligne de test")
    @excel.save
    assert File.file?(@excel.path)
    assert File.unlink(@excel.path)         # Suppression du fichier
  end
  
  ##########################################
  # Test de lecture des données sur le fichier Excel
  ##########################################
  
  def test_title_1_writes_correct_values
    text   = 'Titre looooooooooong'
    # Valeur attendue
    values = [[text]]
    assert_write_correct_value_with_function(values, 'title_1', text)
  end

  def test_title_2_writes_correct_values
    text   = 'Titre looooooooooong'
    values = [[text]]
    assert_write_correct_value_with_function(values, 'title_2', text)
  end
  
  def test_title_3_writes_correct_values
    text   = 'Titre looooooooooong'
    # 'nil' vient du fait qu'un titre de niveau 3 commence sur la seconde colone
    values = [[nil, text]]
    assert_write_correct_value_with_function(values, 'title_3', text)
  end
  
  def test_line_writes_correct_values
    values = [[nil, 'Valeurs', 'sur', 'une', 'ligne']]
    assert_write_correct_value_with_function(values, 'line', *['Valeurs', 'sur', 'une', 'ligne'])
  end
  
  def test_h_table_writes_correct_values
    # Valeurs attendues
    values = [
        [nil, 'colone 1', 'colone 2', 'colone 3'],
        [nil],
        [nil, 'a', 'b', 'c'],
        [nil, nil, nil, 'd'],
        [nil, 'e'],
        [nil, nil, 'f']
      ]
    # Ce que j'écris dans le fichier
    headers = ['colone 1', 'colone 2', 'colone 3']
    write   = [ ['a', 'b', 'c'], [nil, nil, 'd'], ['e'], [nil, 'f', nil] ]
    assert_write_correct_value_with_function(values, 'h_table', write, headers, {:border_bottom => true})
  end  
  
  # Il n'y a pas de test où h_table est appelé avec :border_bottom => false
  # car la ligne de démarquation est ajoutée dans tous les cas
  
  def test_v_table_writes_correct_values
    values = [
        [nil, 'colone 1', 'valeur 1'],
        [nil, 'colone 2', 'valeur 2'],
        [nil, 'colone 3'            ],
        [nil,        nil, 'valeur 3'],
        [nil, 'colone 4', 'valeur 4']
      ]
    headers = ['colone 1', 'colone 2', 'colone 3', nil, 'colone 4']
    write   = ['valeur 1', 'valeur 2', nil, 'valeur 3', 'valeur 4']
    assert_write_correct_value_with_function(values, 'v_table', write, headers)
  end 
  
  def test_complete_write_document
    @excel.title_1('Titre 1')
    @excel.title_2('Titre 2')
    @excel.title_3('Titre 3')
    @excel.line("Ligne de test")
    @excel.h_table([%w{a b c}, %w{d e f}, %w{g h}], ['toto', 'tata', 'tutu'], {:border_bottom => true})
    @excel.v_table(%w{toto tata tutu}, ['colone 1', 'colone 2', 'colone 3', 'colone 4'])
    @excel.newline
    @excel.line("Autre ligne de test")
    @excel.save
    
    assert File.file?(@excel.path)
    
    expected_values  = []
    # Les 'nil' sont dus au fait que la première colone est vide
    expected_values << ["Titre 1"]
    expected_values << ["Titre 2"]
    expected_values << [nil, "Titre 3"]
    expected_values << [nil, "Ligne de test"]
    expected_values << [nil, 'toto', 'tata', 'tutu']
    expected_values << [nil]
    expected_values << [nil, 'a', 'b', 'c']
    expected_values << [nil, 'd', 'e', 'f']
    expected_values << [nil, 'g', 'h']
    expected_values << [nil, 'colone 1', 'toto']
    expected_values << [nil, 'colone 2', 'tata']
    expected_values << [nil, 'colone 3', 'tutu']
    expected_values << [nil, 'colone 4']
    expected_values << [nil]
    expected_values << [nil, "Autre ligne de test"]
    
    assert_equal expected_values, array_from_excel(@excel.path)
    
    assert File.unlink(@excel.path)
    
    assert ! File.file?(@excel.path)    
  end
  
  
  ##########################################
  # Méthodes privées définies pour les tests
  ##########################################
  
  private

  def assert_method_returns_self(method, *args)
    return_class = @excel.send(method, *args).class
    
    assert_equal @excel.class, return_class
  end

  def assert_method_increments_current_row_index(number_of_lines_added, method, *args)
    row_index      = @excel.current_row_index
    index_expected = row_index + number_of_lines_added

    @excel.send(method, *args)

    assert_equal index_expected, @excel.current_row_index
  end
  
  def assert_write_correct_value_with_function(expected_values, method, *args)
    @excel.send(method, *args)
    @excel.save
    
    assert       File.file?(@excel.path)
    
    assert_equal expected_values, array_from_excel(@excel.path)
    
    assert       File.unlink(@excel.path)
    
    assert       ! File.file?(@excel.path)
  end
  
  def upload_excel_file(filepath)
    bytes    = open(filepath,'rb'){|io| io.read}
    excel    = SpreadsheetToArray::FromExcel.new(bytes)
  end
  
  def array_from_excel(filepath)
    excel = upload_excel_file(filepath)
    excel.values
  end  
  
end
