require 'rake'

task :build_gem do
  File.open("#{Dir.pwd}/batch_jobs/build_doc.bat", 'w') {|file|
    file.write("cd #{Dir.pwd}\n")
    file.write("rdoc --op #{Dir.pwd}/doc\n")
  }
  system("#{Dir.pwd}/batch_jobs/build_doc.bat")
  FileUtils.rm_rf(Dir.glob("batch_jobs/build_doc.bat"))
  system("#{Dir.pwd}/batch_jobs/build_gem.bat")
  system("#{Dir.pwd}/batch_jobs/install_gem.bat")

  gem_files = Dir.glob("rub_a_scripts*.gem")
  FileUtils.rm_rf(gem_files)
end