  
namespace :summarize_data do
  #desc "Run the matching of actual PVS Sites"
  #task :final_matched => :environment do
  #  puts "Updating PVC Final Sites..."
  #  FinalMatched.update_analysis
  #  puts "Finished"
  #end
  #
  #desc "Run the matching of actual PVS Sites"
  #task :final_matched => :environment do
  #  puts "Updating PVC Final Sites..."
  #  FinalMatched.update_analysis
  #  puts "Finished"
  #end
  
  #desc "Auto pull for Rates"
  #task :dash => :environment do
  #  puts "Pulling.."
  #  Dash.update_analysis
  #  puts "Finished"
  #end
  
  
  desc "Pull all data for report and run report"
  task :excel_report => :environment do
    puts "STANDBY pulling IMF rates"
    ImfDatum.update_analysis
    puts "OK finished pulling rates"
    
    puts "STANDBY pullling Azerbaijan rates"
    AzerbaijanPull.update_analysis
    puts "OK Azerbaijan rates pulled"
    
    puts "STANDBY pulling census numbers"
    ExportVol.update_analysis
    puts "OK census done"
    
    puts "Creating report..."
    ExcelReport.update_analysis
    puts "Nice"
  end
    
  #desc "Create Rate Exchange Excel Report"
  #task :excel => :environment do
  #  puts "creating.."
  #  ExcelReport.update_analysis
  #  puts "Finished"
  #end
  #
  #desc "Auto Pull From IMF for Rates"
  #task :rate_exchange => :environment do
  #  puts "Pulling.."
  #  ImfDatum.update_analysis
  #  puts "Finished"
  #end
  #
  #
  #desc "Auto Pull Export Volume from Census"
  #task :export_vol => :environment do
  #  puts "Pulling.."
  #  ExportVol.update_analysis
  #  puts "Finished"
  #end
  #
  #desc "Azerbaijan"
  #task :az => :environment do
  #  puts "Pulling.."
  #  AzerbaijanPull.update_analysis
  #  puts "Finished"
  #end
  
end