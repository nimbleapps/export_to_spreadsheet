# Adding lib to the load path
$:.push File.expand_path("../lib", __FILE__)

Gem::Specification.new do |s|
  s.name        = 'export_to_spreadsheet'
  s.version     = '0.2.0'
  s.platform    = Gem::Platform::RUBY
  s.summary     = "Export data to Google or Excel"
  s.description = "Export model data to Google Spreadsheets or Excel through Apache POI"
  s.homepage    = "http://github.com/nimbleapps/export-to-spreadsheet"
  s.authors     = ["Nimble Apps"]
  s.email       = "dev@nimble-apps.com"
  s.files       = Dir["lib/**/*"]
  s.require_paths = ["lib"]

  # Use Class#subclasses (used in SalesClicExporter::Document#allowed_extensions) is defined by Rails 3
  # It also exists in Rails 2 but the spec is different.
  s.add_dependency 'rails', '>= 3.0'

  # Plug-in that handles Google Spreadsheets API communication
  s.add_dependency 'google-spreadsheet-ruby', '>=0.1.5'

  s.add_dependency 'rjb'

  # Plug-in that handles Google API authentication
  s.add_dependency 'oauth'

  s.add_dependency 'oauth-plugin', ">=0.4.1"

  s.add_development_dependency 'yard'
  s.add_development_dependency 'highline'
  s.add_development_dependency 'sqlite3'
  s.add_development_dependency 'ruby-debug'
end
