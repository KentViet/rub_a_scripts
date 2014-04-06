Gem:
Rub_A_Scripts

Description:

An extension of RubyInline gem. rub_a_scripts allows you to write VBScript within
Ruby code. The extensions are then automatically loaded into the class/module that defines it.
Allowing ruby to call on VBScript function and procedure. Basic data types are mapped as such
facilitate the exchange of data between the two languages.

Example:

class MyVBScriptClass                   #==> this is Ruby.
  inline :VBScript do |builder|

    builder.vbscript '
    Function fnSum(var1,var2)           #==> this is VBScript.
      vrSum = var1 + var2
      fnSum = vrSum
    End Function
    '

    builder.vbscript '
    Function fnReturn2DArray()          #==> return to ruby an array [[1,3],["Hello","World"],true]
      array1 = Array(1,3)
      array2 = Array("Hello", "World")
      fnReturn2DArray = Array(array1,array2,True)
    End Function
    '

  end
end

@vb_script = MyVBScriptClass.new()
puts @vb_script.fnSum(3,5)              #==> Ruby passing 3,8 to VBScript, will return 8.
puts @vb_script.fnReturn2DArray()       #==> will return [[1,3],["Hello","World"],true].

A very simple examples. But Essentially you can load any VBScript you want. Basic data types are mapped so you can send
and retrieve whatever you need(except for object - you'll get nil for it cause I don't know how to map complex object).
For more examples see the unit tests.

How:

Upon instantiation of the class. Ruby will gather all the VBScript source code and then create a .vbs file in the
RubyInline folder. It will also add a VBScript method to that .vbs file that is responsible for picking up arguments
send to the .vbs file. Then later on send the output to stdout so that Ruby can pick up the output and then
format it as the result.
    - For examples, in the above example a file will be created in the RubyInline folder: VBScriptXXXXXXXXX.vbs
    - Then Ruby will execute that file: cscript.exe VBScriptXXXXXXXXX.vbs fnSum(3,5) \\NOLOGO
    - There is a method inside VBScriptXXXXXXXXX.vbs named fxWriteOutput which will take fnSum(3,5) and do an Eval on it.
    - It will take the result and then put it to the stdout.
    - Since Ruby executed the system process that run VBScriptXXXXXXXXX.vbs it will collect the last line in the stdout.
    - Covert that to the appropriate type and then reassign it to the Ruby function call.

Warning:
    - after you instantiate the class you might want to use .delete_file to remove the VBScript file.
    - you can use clean_cache to all VBScript file in the RubyInline folder

Q&A:
    1. Why name it Rub_A_Scripts?
        - because (Ruby + VBScript). Its like a marriage, both have to give
        up something to make it works. :)
    2. But why? VBScript is platform dependent, slower, less elegant than Ruby. Anything that can
        be done in VBScript can easily be done in Ruby.
        - There are a lot of existing VBScript and build scripts that we can reuse or integrated
        into Ruby.
        - VBScript have access to various tools/api that by giving Ruby access to VBScript we
        can then tap into those tools/api.

License:
The MIT License (MIT)

Copyright (c) 2014 <Kent T. Le>

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.