namespace :cron do
  desc "Delete 1 day old excel files from /tmp/"
  task :tidy_exported_tmp_files do
    dir           = Rails.root.to_s + '/tmp'
    deleted_files = []
    
    Dir.entries(dir).each do |filename|
      filepath = dir.to_s + '/' + filename      
      
      # Suppression des fichiers d'export Excel vieux de 24h
      if File.extname(filepath) == '.xls' && File.ftype(filepath) == 'file' && File.mtime(filepath) < (Time.now - 24*60*60)
        File.unlink(filepath)
        deleted_files << filename
      end
    end
    
    if deleted_files.empty?
      puts "Sorry, we did not find any files to delete"
    else
      puts "#{deleted_files.size} files have been deleted: #{deleted_files.join(', ')}"
    end
    
  end
end