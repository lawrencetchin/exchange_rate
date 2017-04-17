class AzerbaijanPull < ApplicationRecord
  def self.update_analysis
    require 'rubygems'
    require 'nokogiri'
    require 'open-uri'
    
    puts "Deleting Azerbaijani Manat current records.."
    ImfDatum.where('currency_name = ?', 'Azerbaijani Manat').delete_all
    puts "Deleted"
    
    date = Date.today
    day = date.mday
    month = date.mon
    year = date.year
    
    page = Nokogiri::HTML(open("https://en.cbar.az/other/azn-rates?act=betweenForm&from%5Bday%5D=1&from%5Bmonth%5D=1&from%5Byear%5D=1995&to%5Bday%5D=#{day}&to%5Bmonth%5D=#{month}&to%5Byear%5D=#{year}&rateID=usd", ssl_verify_mode: OpenSSL::SSL::VERIFY_NONE))   
      
    #dates = page.css("div[class='betweenResults']").css('td[class="date"]').xpath
    #rates = page.css("div[class='betweenResults']").css('td[class="rate"]').text
    #full = page.css("div[class='betweenResults']").css('tr').xpath
    
    final =  page.css("div[class='betweenResults']").css('tr')
    
    final_arr = []
    final.each do |trial|
      row = trial.css('td')
      #puts row.text
      final_arr << [Date.parse(row.text[0..9]), row.text[10..-1]]
      #puts "'#{final_arr}'"
      #row.each do |td|
      #  puts td
      #end
    end
    
    puts "Found #{final_arr.length} records"
    ActiveRecord::Base.transaction do
          puts "saving new records..."
          final_arr.each do |f|
            ActiveRecord::Base.connection.execute "INSERT INTO imf_data (date, currency_name, rate) VALUES ('#{f[0]}', 'Azerbaijani Manat', #{f[1]})"
          end
      end 
    
  end
end
