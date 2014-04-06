require 'rub_a_scripts'
require 'minitest'
require 'minitest/autorun'

class MyVBScriptClass

  inline :VBScript do |builder|
    builder.vbscript '
    Function fnSum(var1,var2)
      vrSum = var1 + var2
      fnSum = vrSum
    End Function
    '

    builder.vbscript '
    Function fnSubtract(var1,var2)
      vrSubtract = var1 - var2
      fnSubtract = vrSubtract
    End Function
    '

    builder.vbscript '
    Function fnString(input)
      fnString = input & " World!!!"
    End Function
    '

    builder.vbscript '
    Function fnBoolean(input)
      if input = True then
        fnBoolean = False
      else
        fnBoolean = True
      End If
    End Function
    '

    builder.vbscript'
     \'WScript.Echo "Hello Cruel Cruel World!!!"
    '

    builder.vbscript '
    Function fnArray(input)
      fnArray = input(2)
    End Function
    '

    builder.vbscript '
    Function fn2Array(input, input2)
      fn2Array = input2(5)
    End Function
    '

    builder.vbscript '
    Function fnNull()
      fnNull = "nothing"
    End Function
    '

    builder.vbscript '
    Function fnDecimal()
      fnDecimal = 123.13
    End Function
    '

    builder.vbscript '
    Function fnDecimalNegative()
      fnDecimalNegative = -1234567.13
    End Function
    '

    builder.vbscript '
    Function fnReturnArray()
     fnReturnArray = Array(3,9,4,0,8,2)
    End Function
    '

    builder.vbscript '
    Function fnReturn2DArray()
      array1 = Array(1,3)
      array2 = Array("Hello", "World")
      fnReturn2DArray = Array(array1,array2,True)
    End Function
    '

    builder.vbscript '
    Function fnProcessMultiDimArray(input)
      fnProcessMultiDimArray = input
    End Function
    '

    builder.vbscript '
    Function fnObject()
      set fso = CreateObject("Scripting.FileSystemObject")
      Set fnObject = fso
    End Function
    '

    builder.vbscript '
    Sub subString()
      WScript.Echo "I am inside a Procedures!!!"
      response.write("I was written by a sub procedure")
    End Sub
    '

  end

end

class TestRubAScript < Minitest::Test
  def setup
    @vb_script_obj = MyVBScriptClass.new()
    @my_str = "Hello"
    @my_array = ['a', 'b', 'c', 'd', 'e', 'f']
    @my_array2 = ["1", "2", "3", "4", "5", "6"]
    @my_array3 = [9,[0],[[1],[[2],[3]],[4]],"Hello World"]
    @error_str = "Error: Object doesn't support this property or method. Sorry!!! Object are not mapped."
  end

  def test_fnSum
    assert_equal 8, @vb_script_obj.fnSum(3,5)
  end

  def test_fnSubtract
    assert_equal 5, @vb_script_obj.fnSubtract(8,3)
  end

  def test_fnSubtractNegative
    assert_equal -5, @vb_script_obj.fnSubtract(3,8)
  end

  def test_fnString
    assert_equal "Hello World!!!", @vb_script_obj.fnString('Hello')
  end

  def test_fnString_with_var
    assert_equal "Hello World!!!", @vb_script_obj.fnString(@my_str)
  end

  def test_fnBoolean_true
    assert_equal false, @vb_script_obj.fnBoolean(true)
  end

  def test_fnBoolean_false
    assert_equal true, @vb_script_obj.fnBoolean(false)
  end

  def test_fnArray
    assert_equal "c", @vb_script_obj.fnArray(@my_array)
  end

  def test_2fnArray
    assert_equal 6, @vb_script_obj.fn2Array(@my_array, @my_array2)
  end

  def test_fnNull
    assert_equal nil, @vb_script_obj.fnNull
  end

  def test_fnDecimal
    assert_equal 123.13, @vb_script_obj.fnDecimal
  end

  def test_fnDecimal
    assert_equal -1234567.13, @vb_script_obj.fnDecimalNegative
  end

  def test_fnReturnArray
    assert_equal [3,9,4,0,8,2], @vb_script_obj.fnReturnArray()
  end

  def test_fnReturn2DArray
    assert_equal [[1,3],['Hello', 'World'], true], @vb_script_obj.fnReturn2DArray()
  end

  def test_fnProcessMultiDimArray
    assert_equal @my_array3, @vb_script_obj.fnProcessMultiDimArray(@my_array3)
  end

  def test_fnObject
    assert_equal @error_str, @vb_script_obj.fnObject
  end

  def test_subString()
    assert_equal nil, @vb_script_obj.subString()
  end

end

Minitest.after_run {
  @test = MyVBScriptClass.new()
  @test.delete_file
  @test.clean_cache
}
