Gem::Specification.new do |s|
  s.name                              = 'rub_a_scripts'
  s.version                           = '0.0.1'
  s.default_executable                = 'rub_a_scripts'
  s.required_rubygems_version         = Gem::Requirement.new(">= 0") if s.respond_to? :required_rubygems_version=
  s.authors                           = 'Kent T. Le'
  s.date                              = '2014-03-25'
  s.description                       = "An extension of RubyInline. Allow for the use of VBScript in Ruby."
  s.email                             = "Le.78@osu.edu"
  s.files                             = ["lib/rub_a_scripts.rb", "bin/rub_a_scripts", 'Rakefile']
  s.test_files                        = ["test/test_rub_a_scripts.rb"]
  s.homepage                          = %q{https://github.com/KentViet/rub_a_scripts}
  s.require_paths                     = ['lib', 'test', 'bin']
  s.rubygems_version                  = %q{1.8.2}
  s.summary                           = %q{An extension of RubyInline gem. rub_a_scripts allows you to write VBScript within
                                           Ruby code. The extensions are then automatically loaded into the class/module that defines it.
                                           Allowing ruby to call on VBScript function and procedure. Basic data types are mapped as such
                                           facilitate the exchange of data between the two languages.}
end