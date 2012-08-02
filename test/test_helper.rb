require 'rubygems'
require 'test/unit'
require 'active_record'
require 'active_support/test_case'

require File.dirname(__FILE__) + '/../init.rb'

# Modifie la classe SalesClicExporter::Document::GoogleSpreadsheets 
# pour permettre d'avoir accès à l'objet 
SalesClicExporter::Document::GoogleSpreadsheets.class_eval do
  def google_doc ; @google ; end
end

class TestObject #< ActiveRecord::Base
  attr_accessor :name
  attr_accessor :value

  include ExportToSpreadsheet
end

def create_obj
  #return TestObject.create!(:name => "Affaire", :value => 999.7)
  test_obj       = TestObject.new
  test_obj.name  = "Affaire"
  test_obj.value = 999.7
  return test_obj
end

# On rend accessible la variable d'instance current_row_index seulement
# pour les tests
SalesClicExporter::Document::Base.class_eval do; attr_accessor :current_row_index; end

###############

def load_excel_file(filename)
  filepath = File.dirname(__FILE__) + '/resources/' + filename
  bytes    = IO::read(filepath)
  #bytes    = File.read(filepath, 'rb') { |fin| fin.bytes.to_a }
  excel    = SpreadsheetToArray::FromExcel.new(bytes)
end

def array_from_excel(filename)
  excel = load_excel_file(filename)
  excel.values
end