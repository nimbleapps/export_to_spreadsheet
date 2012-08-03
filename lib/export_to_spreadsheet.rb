# This is the module to include into an ActiveRecord model in order to make its data exportable.
# You will have to also define a compose_export method that tells the plug-in what data to export.
# @example
#   class Foo < ActiveRecord::Base
#     include ExportToSpreadsheet
#
#     def compose_export
#       @doc.write do |d|
#         d.title_1('Export of foo')
#         d.line("My foo is #{self.name}")
#       end
#     end
module ExportToSpreadsheet
  # If we have a file that tweaks the export plug-in from the point of view of the application, we require it
  if defined?(Rails)
    required_file = "#{Rails.root}/lib/before_prepare_export.rb"
    require required_file if File.exists?(required_file)
  end

  # Whether we want to use fork to run the export or not
  @@export_use_fork = true
  
  def self.export_use_fork= val
    @@export_use_fork = val
  end

  # Generates a complex String to ensure that a filename is uniq
  def make_token
    # From the restful-authentication plug-in
    args = [ Time.now, (1..10).map{ rand.to_s } ]
    Digest::SHA1.hexdigest(args.flatten.join('--'))
  end

  # The actual export (starting the JVM and exporting) is done in a forked process that can be terminated once the
  # file is written.
  # This is to make sure that the (high amount of) memory used by the JVM is freed once its job is over.
  # With MySQL this messes a little with database connections so we have to close the connection and reopen it in each process.
  # On Windows, this does not work so when do a basic export then.
  #
  # @param [Hash] args hash for export configuration
  # @option args [String] :filename        filename
  # @option args [String] :user_team_id    the id of the user that performs the export
  # @option args [String] :export_type     the id of the user that performs the export
  # @option args [String] :pipeline_date   the date that is used to view the pipeline
  # @return the document it-self, which is a class of {SalesClicExporter::Document}
  # @raise [RunTimeError] an exception is raised when an error occurs during the export. The log has to be checked to see what happened
  def to_excel(*args)
    # Setting up a few variable so we can find the file once it has been written
    export = prepare_to_excel(*args)

    # Closing the DB connection before forking, we store the config to be able to reconnect easily
    dbconfig = ActiveRecord::Base.remove_connection

    # fork forks the process and returns
    #   the pid of the child process in the father process
    #   and the nil in the child process
    # so we know in which process we are depending on the returned value
    begin
      if @@export_use_fork
        if fork
          # This is the father process
          ActiveRecord::Base.establish_connection(dbconfig)
          # We just wait for the child process to have finished the export
          # wait2 returns [pid, exit_code] # exit_code has the class Process::Status
          exit_code = Process.wait2.last.to_i
        else
          begin
            # This is the child process
            ActiveRecord::Base.establish_connection(dbconfig)
            # Doing the actual export
            # This is what starts the JVM
            compute_to_excel(*args)
            exit!(0) # Terminating the process
          rescue Exception => e # need to make sure the child process finishes even if the export fails
            logger.error e.message
            logger.error e.backtrace.join("\n")
            exit!(1)
          end
        end
      else
        raise "not even trying"
      end

    # Handles at least the Windows case where the fork function is not implemented.
    # In this case do a simple export without forking the process
    rescue Exception => e
      ActiveRecord::Base.establish_connection(dbconfig)
      compute_to_excel(*args)
      exit_code = 0
    end

    # If the export did not finish normally, we raise an error
    raise "Unexpected error during export" if exit_code != 0

    # It worked, we return the export
    return export
  end

  # @overload to_google_spreadsheets(oauth_token)
  #   Exports a model to a Google Spreadsheet doc using a OAuth token
  # @overload to_google_spreadsheets(google_login, google_password)
  #   Exports a model to a Google Spreadsheet doc using a Google account credentials
  def to_google_spreadsheets(*args)
    prepare_export(:google_spreadsheets, *args)
    return compute_export(:google_spreadsheets, *args)
  end

  private

  # We had to split the excel export in 2 parts to have the file named first
  # and have the JVM started in a second process
  # This just handles naming
  def prepare_to_excel(*args)
    return prepare_export(:excel, *args)
  end

  # This does the heavy lifting
  def compute_to_excel(*args)
    return compute_export(:excel, *args)
  end

  # Configures export
  def prepare_export(target, options = {})

    # #before_prepare_export_first is defined in /lib/before_prepare_export.rb
    if self.respond_to? :before_prepare_export_first
      options.merge!(before_prepare_export_first(target, options))
    end

    default_filename = "Export of #{self.class} #{Date.today.strftime("%Y-%m-%d")}"
    filename         =
      (options && ! options.empty? && options[:filename]) ? options.delete(:filename) : default_filename
    
    # If we need customization for a specific model, #setup_export method returns a Hash
    # that will impact the generated document
    if self.respond_to?(:setup_export) && setup_export.is_a?(Hash)
      options.merge!(setup_export)
    end
    
    @doc =
      case target
        when :excel
          # Le premier nom correspond à celui que verra le client, le second
          # est une clé générée à la volée pour créer un fichier unique sur le serveur
          SalesClicExporter::Document::Excel.new(filename, make_token)
        when :google_spreadsheets
          SalesClicExporter::Document::GoogleSpreadsheets.new(filename, options)
        else
          raise
      end

    return @doc
  end

  # Last step of an export.
  def compute_export target, options = {}
    # In the excel case, we need to go further in the initialize process
    @doc.finish_initialize(options) if target == :excel

    # Calls the methods that defines the data to export, inside a model
    compose_export

    return @doc
  end

end