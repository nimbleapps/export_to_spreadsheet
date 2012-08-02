require 'bundler'
Bundler::GemHelper.install_tasks

require 'rake/testtask'
require 'yard'
desc 'Generate documentation for the export_to_files plugin using yard.'
YARD::Rake::YardocTask.new do |t|
  t.files   = ['lib/**/*.rb', 'README', 'CHANGELOG']
end

Rake::TestTask.new(:test) do |t|
  t.libs << 'lib'
  t.libs << 'test'
  t.pattern = 'test/**/*_test.rb'
  t.verbose = false
end
