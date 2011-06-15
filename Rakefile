require 'rubygems'  
require 'rake'  
require 'rake/testtask'
require 'yard'
require 'echoe'

Echoe.new('export_to_spreadsheet') do |p|  
  # The version can be specified as a 2nd argument after the name, or completly left (in this case
  # the last version in the CHANGELOG file will be used)
  p.version         =  '0.1.2'
  p.description     = "Export model data to Google Spreadsheets or Excel through Apache POI"  
  p.url             = "http://github.com/nimbleapps/export-to-spreadsheet"  
  p.author          = "Nimble Apps"  
  p.email           = "dev@nimble-apps.com"  
  p.ignore_pattern  = ["tmp/*", "script/*"]

  # 'portablecontacts' is in the dependencies because the most recent version of 'oauth-plugin' available
  # on github (v0.3.14) has an implicit dependency against this gem (in 'google_token.rb').
  # This is corrected by this commit https://github.com/pelle/oauth-plugin/commit/3ff9ccf444f2810214ada6fca3da34008ff52795
  # but there is no release yet with this commit.
  p.runtime_dependencies = ['portablecontacts', 'google-spreadsheet-ruby >=0.1.5', 'rjb', 'oauth', 'oauth-plugin']

  p.development_dependencies = ['yard', 'highline', 'active_support']
end

desc 'Generate documentation for the export_to_files plugin using yard.'
YARD::Rake::YardocTask.new do |t|
  t.files   = ['lib/**/*.rb', 'README', 'CHANGELOG']
end
