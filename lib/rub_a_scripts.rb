require "inline"
require 'Date'
module Inline

  ##
  # VBScript Inline - allow simple inline of VBScript within Ruby.
  #
  class VBScript

    ##
    # extra_args: Allow the user to specify any other arguments they want to add. By default //Nologo is added.
    #
    attr_accessor :extra_args

    def initialize(mod)
      @context = mod
      @src = ""
      @imports = []
      @sigs = []
      @file_name = ''
      @extra_args = '//Nologo'
    end

    ##
    # Use by RubyInline to designate new file or to reuse cache. For VBScript
    #   we simply create a new file every times so this will always be set to false.
    #   Also crapped in here a method to clear out the inline folder if files are
    #   more than 3 days old.
    #
    def load_cache
      clear_inline_folder
      false
    end

    ##
    # This method will clear out the inline directory if any files is more than 3 days old.
    #
    def clear_inline_folder
      files = Dir.glob("#{Inline.directory}/VBScript*")
      files.each do |file|
        file_time = File.ctime(file)
        today_date = Date.today
        if ((file_time.day - today_date.day) >= 2) and (file_time.month >= today_date.month) and (file_time.year >= today_date.year)
          FileUtils.rm_rf(file)
        end
      end
    end

    ##
    # Parse your VBScript to find a function name. Store them sor Ruby
    #    can create a Ruby method with the exact name and parameters list.
    #
    def vbscript(src)
      @src << src << "\n"
      signature = src.match(/Function\W+([a-zA-Z0-9_]+)\((.*)\)/)
      @sigs << [signature[1], signature[2], signature[3]] unless !signature
      if !signature
        signature = src.match(/Sub\W+([a-zA-Z0-9_]+)\((.*)\)/)
        @sigs << [signature[1], signature[2], signature[3]] unless !signature
      end
    end

    ##
    # @build. Build source VBScript file to be later used for execution.
    #   - grabbing all your vbscript source code
    #   - creating a vbscript file and then inputing in the code
    #   - it will also add an extra vbscript function that will
    #   later be called to pass in parameters and output result to
    #   a file so Ruby can pick up.
    #
    def build
      @load_name = @name = "VBScript#{@src.hash.abs}"
      @file_name = "#{Inline.directory}/#{@name}.vbs"
      @src << extra_input_fx << "\n"
      File.open(@file_name, 'w') {|file| file.write(@src)}
    end

    ##
    # This method is used to dynamically create all the Ruby methods
    #   that is associate with their VBScript methods. No doubt the
    #   most important method of this module as it's reponsible for
    #   creating handling interactions between Ruby and VBScript.
    #
    def load
      @context.module_eval "const_set :#{@name}, Inline::VBScript;"

      @context.module_eval "
      def clean_cache
        files = Dir.glob(\"\#{Inline.directory}/VBScript*\")
        files.each {|file| FileUtils.rm_rf(file)}
      end
      def delete_file
        File.delete('#{@file_name}') if File.exist?('#{@file_name}');
      end
      "

      @context.module_eval "
      def construct_arguments(*args)
        @argument_to_send = ''
        if args.class == Array
          args.each_with_index do |top_level_item, args_index|
            @argument_to_send += \",\" if args_index > 0
            if top_level_item.class == Array
              @argument_to_send += \"Array(\"
              top_level_item.each_with_index do |inner_item, top_level_index|
                @argument_to_send += \",\" if top_level_item.count != top_level_index and top_level_index > 0
                @argument_to_send += construct_arguments(inner_item)
              end
              @argument_to_send += \")\"
            else
              case
                when top_level_item.class == String
                  @argument_to_send += %q[\"] + %Q[\#{top_level_item.to_s}] + %q[\"]
                when top_level_item.class == Fixnum
                  @argument_to_send += top_level_item.to_s
                when top_level_item.class == TrueClass
                  @argument_to_send += \"True\"
                when top_level_item.class == FalseClass
                  @argument_to_send += \"False\"
                else
                  @argument_to_send += top_level_item.to_s
              end
            end
          end
        end
        return @argument_to_send
      end
      "

      @sigs.each do |sig|
        @context.module_eval "
        def #{sig[0]}(*args);
          @arguments_to_send = construct_arguments(args)
          @arguments_to_send.sub!(\"Array(\",\"\").chomp!(\")\")
          @arguments_to_send.gsub!(/\"/, \'\\'\')
          @arguments_to_send.chomp!
          execution_str = 'cscript #{@file_name} \"#{sig[0]}' + '(' + @arguments_to_send + ')\" #{@extra_args}';

          IO.popen(execution_str).each do |line|
            case
              when (line.strip == \"nothing\" or line.strip == \"null\" or line.strip == \"Empty\")
                @result = nil
              when line.strip == \"False\"
                @result = false
              when line.strip == \"True\"
                @result = true
              when line.strip[/^[-+]?[0-9]+$/]
                @result = line.strip.to_i
              when line.strip[/^[-+]?[\\d]+(\\.[\\d]+){0,1}$/]
                @result = line.strip.to_f
              when line.strip[/Array\\[.*\\]/]
                begin
                  @result = eval(line.strip.gsub(\"Array\",\"\").gsub(\"True\",\"true\").gsub(\"False\",\"false\"))
                rescue => e
                  @result = line.strip;
                  warn(e)
                end
              else
                @result = line.strip;
            end
          end.close
          return @result
        end"
      end
    end

    ##
    # The inside man. This source code and methods is added and inject into the VBScript.
    #   It's responsible for picking up arguments and executing the VBScript function
    #   being called. Will also write result to stdout so that Ruby can pick up.
    #   A few basic mapping and error handling is also being done here vs the load method on the Ruby side.
    #
    def extra_input_fx()
      my_str1 = '
      Function fxWriteOutput
		    InputArgs = WScript.Arguments.Item(0)
        InputArgs = Replace(InputArgs, "\'", chr(34))
	      if WScript.Arguments.Count() > 0 Then
          On Error Resume Next
            fx_value = Eval(InputArgs)
            If Err.Number <> 0 Then
              If Err.Description = "Object doesn\'t support this property or method" Then
                fx_value = "Error: " & Err.Description & ". Sorry!!! Object are not mapped."
              End If
              Err.Clear
            End If
          On Error Goto 0
          ' + "\n"
      my_str_3 = '
            if IsArray(fx_value) = False Then
              if fx_value = True Then
                WScript.Echo "True"
              elseif IsEmpty(fx_value) Then
                WScript.Echo "Empty"
              elseif fx_value = False Then
                WScript.Echo "False"
              else
                WScript.Echo fx_value
              End If
            else
              my_array_list = "Array" & abcExpandArray(fx_value)
              WScript.Echo my_array_list
            End if
        End if
      End Function
      fxWriteOutput()

      Function abcExpandArray(fx_value)
	      If IsArray(fx_value) Then
          For i = (lbound(fx_value)) to ubound(fx_value)
            if i > 0 Then
              finalStr = finalStr & "," & abcExpandArray(fx_value(i))
            Else
              finalStr = finalStr & abcExpandArray(fx_value(i))
            End If
          Next
          abcExpandArray = "[" & finalStr  & "]"
        Else
          If IsNumeric(fx_value) Then
            abcExpandArray = fx_value
          ElseIf fx_value = True Then
            abcExpandArray =  "True"
          ElseIf IsEmpty(fx_value) Then
            abcExpandArray =  "Empty"
          ElseIf fx_value = False Then
            abcExpandArray =  "False"
          Else
            abcExpandArray = """" & fx_value & """"
          End If
	      End If
      End Function
      '
      return (my_str1 + my_str_3)
    end

  end

end