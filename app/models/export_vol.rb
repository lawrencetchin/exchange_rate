class ExportVol < ApplicationRecord
  def self.update_analysis
    require 'rubygems'
    require 'nokogiri'
    require 'open-uri'
    
      
      puts "Deleting All Records before New Pull...."
      ExportVol.delete_all()
      
      country_codes = ["c6021", "c3510", "c1220", "c5700", "c4279",
                       "c4280", "c4120", "c4840", "c5820", "c5330",
                       "c4759", "c5880", "c2010", "c4210", "c6141",
                       "c4039", "c5650", "c4621", "c5590", "c5800",
                       "c4700", "c4010", "c5490", "c4632"]
    	#Vietnam EXCLUDED FOR NOW     
      countries = [ "Australia", "Brazil", "Canada", "China",	"France",
									   "Germany","Great Britain","Greece", "Hong Kong", "India",
									   "Italy", "Japan",	"Mexico",	"Netherlands", "New Zealand",
									   "Norway", "Philippines", "Russia", "Singapore", "South Korea",
									   "Spain", "Sweden", "Thailand", "Azerbaijan"]
                    
      country_codes = country_codes.zip(countries)
      
      country_codes.each do |t|
        census = Nokogiri::HTML(open("https://www.census.gov/foreign-trade/balance/#{t[0]}.html", ssl_verify_mode: OpenSSL::SSL::VERIFY_NONE))
        census_data = census.css("div[id='middle-column']").css('table').css('td').text
        
        census_data = census_data.split(" ")
        census_data = census_data.in_groups_of(5)
        census_data.slice!(-1)
        
        census_data.map! {|c| c.take(3)}
        census_data.map {|d| d.push(t[1])}
        #puts "CENSUS #{census_data}"
        
        ActiveRecord::Base.transaction do
          puts "SAVING ALL RECORDS FOR #{t[1]}.."
           census_data.each do |f|
              if f[0] != "TOTAL"
               f[2].gsub!(',','')
               ActiveRecord::Base.connection.execute "INSERT INTO export_vols (month, year, export_vol, country) VALUES ('#{f[0]}', '#{f[1]}', #{f[2].to_f}, '#{f[3]}')"
              end
            end
        end  
      end
      
  end
end
