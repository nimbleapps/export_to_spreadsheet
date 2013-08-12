# -*- encoding : utf-8 -*-
require 'test_helper'

class SpreadsheetToArrayTest < ActiveSupport::TestCase
  
  def test_values_are_read_correctly
    assert_nothing_raised do
      
      array = array_from_excel('valid_excel_import.xls')
      
      first_line_values       = ['Objectif', 5936.535, 7730.663, 9524.791 , 11318.919, 13113.047, 14907.175, 16701.303, 18495.431]
      second_line_values      = ['Réalisé', 4635.394, 9695.558, 14755.722, 19815.886, 24876.050, 29936.214, 34996.378, 40056.542]     
      

      excel_first_line_values = array[2]
      
      excel_sec_line_values   = array[3]
      
      assert_equal  first_line_values[0] , excel_first_line_values[0]
      assert_equal  second_line_values[1], excel_sec_line_values[1]

      # assert_in_delta used because there is a problem with float precision when reading from xls files
      (1..first_line_values.size-1).each do |n|
        assert_in_delta  first_line_values[n] , excel_first_line_values[n], 0.0001
      end
      (1..second_line_values.size-1).each do |n|
        assert_in_delta  second_line_values[n], excel_sec_line_values[n]  , 0.0001
      end
    end
  end
  
  def test_exception_is_thrown_when_byte_array_entry_is_not_valid
    assert_raise RuntimeError do
      array_from_excel('invalid_file_format.txt')
    end    
  end
  
end