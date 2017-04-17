class ImfDatum < ApplicationRecord
  def self.update_analysis
    require 'rubygems'
    require 'nokogiri'
    require 'open-uri'

    
    
    
    #####UNCOMMENT FOR PARTIAL DATA PULL
    #start_date = Date.new(2017,1,25) #Date.today.beginning_of_month # your start
    #end_date =  (Date.today)-1.day
    ##USE THIS END_DATE FORMAT FOR SPECIFIC DATE PULL Date.new(2017,1,25)
    #####

    #####UNCOMMENT FOR FULL DATA PULL OF ALL YEARS      
          ###UNCOMMENT FOR FULL DATA PULL FROM BEGINNING (1995)
          #original_start = Date.new(1995,1,31)
          #rolling_date = original_start.end_of_month
          #puts "deleting ALL YEARS AND DATA....."
          #ImfDatum.delete_all()
          #puts "I can't believe you've done this...."
          
          
          ################################################################################
          ################################################################################
          #####################ALTER THIS PORTION ONLY####################################
          ################################################################################
          ################################################################################
          
          ##UPDATE THE DATE##
          #LAST RUN 4/5
          ###################
          
          #Change the start date to the beginning of the month the last time the program was run
          start = Date.new(2017,1,1)
          
          rolling_date = start.end_of_month
          puts "deleting data starting from #{start.beginning_of_month} onwards......"
          ImfDatum.where('date >= ?', start).delete_all
          puts "say bye, because it's gone now.."
          ################################################################################
          ################################################################################
          ################################################################################
          ################################################################################
          ######################ALTER THIS PORTION ONLY###################################
          ################################################################################
          ################################################################################
    #####      
      
      
      #############VIETNAM WEBSITE DATA PULL (INCOMPLETE)
      #current_year = Date.today.year
      #vietnam = Nokogiri::HTML(open("http://www.likeforex.com/misc/historical-rates.php?f=USD&t=VND&y=2016&page=1"))
      ##vietnam_data = vietnam.css("div[class='w7 fr']").css('table[class="panel"]').css('table[width="100%"]').css('td').text
      ##vietnam_data = vietnam_data.split(", #{current_year}")
      #
      #
      #vietnam_date = vietnam.xpath('//table[@class="panel"]/tr/td[2]/text()')
      #vietnam_date.each {|e| e = e.inner_html }
      #vietnam_rate = vietnam.xpath('//table[@class="panel"]/tr/td[3]/text()')
      #vietnam_rate.each {|e| e = e.inner_html }
      #vietnam_data = vietnam_date.zip(vietnam_rate)
      #puts "VIETNAM #{vietnam_data}"
      
      
############IMF WEBSITE DATA PULL  

############UNCOMMENF FOR FULL DATA PULL      
    while rolling_date < Date.today.end_of_month
        puts "current period #{rolling_date}"
        
        my_days = [1,2,3,4,5] # day of the week in 0-6. Sunday is day-of-week 0; Saturday is day-of-week 6.
        result = (rolling_date.beginning_of_month..rolling_date).to_a.select {|k| my_days.include?(k.wday)}
        result = result.in_groups_of(1)
        #puts "Results #{result}"
        page = Nokogiri::HTML(open("https://www.imf.org/external/np/fin/data/rms_mth.aspx?SelectDate=#{rolling_date}&reportType=REP"))
############


        
####        UNCOMMENT FOR PARTIAL DATA PULL
#        ##DELETE ALL RECORDS FROM CURRENT MONTH
#        puts "Deleting current records from #{start_date} up to #{end_date}.."
#        ImfDatum.where('date >= ? and date <= ?', start_date, end_date).delete_all
#        puts "Deleted"
#        
#        puts "Pulling current data for month up to #{end_date}......"
#        
#        my_days = [1,2,3,4,5] # day of the week in 0-6. Sunday is day-of-week 0; Saturday is day-of-week 6.
#        result = (start_date..end_date).to_a.select {|k| my_days.include?(k.wday)}
#        result = result.in_groups_of(1)
#        page = Nokogiri::HTML(open("http://www.imf.org/external/np/fin/data/rms_mth.aspx?SelectDate=#{end_date}&reportType=REP"))   
#      
####     
      
        #Grab the currency data from specific HTML elements
        headers_first = page.css("div[class='fancy']").css('table').first.css("th[class='color3']")
        headers_second = page.css("div[class='fancy']").css('table:nth-child(3)').css("th[class='color3']")
        head_first = []
        head_second = []
        
        headers_first.each do |h|
          head = h.text
          head = head.strip!
          head_first << head
        end
        
        headers_second.each do |h|
          head = h.text
          head = head.strip!
          head_second << head
        end
        

        first_temp = []
        first = page.css("div[class='fancy']").css('table').first.css('tr')
        first.each do |t|
          row = t.css('td').text
          row = row.split(" ")
          temp = []
          name = []
          row.each_with_index do |r,i|
            r = r.gsub(/[\s,]/ ,"")
            if r =~ /\A[-+]?[0-9]*\.?[0-9]+\Z/ 
              temp << r.to_f
            elsif r != 'NA' && r != '(1)' && r != '(2)'
              name << r
            elsif r != '(1)' && r != '(2)'
              temp << r
            end
            
            if i == row.length-1
              name = name.join(" ")
              temp.unshift(name).flatten
              row = temp
              first_temp << row
              temp = []
              name = []
            end
          end
        end
        
          final_first = []
          head_first.each_with_index do |h, ind|
            #puts "current h #{h}"
            first_temp.each_with_index do |f, i|
              #puts "current temp #{f}"
              if ind != 0
                temp = [f[0], h, f[ind]]
                #puts "temp #{temp}"
                final_first << temp
              end
            end
          end
        
        
        
        
        
         
        second_temp = []
        second = page.css("div[class='fancy']").css('table:nth-child(3)').css('tr')
        second.each do |t|
          row = t.css('td').text
          row = row.split(" ")
          #puts "row before #{row}"
          temp = []
          name = []
          #puts "row first #{row}"
          row.each_with_index do |r,i|
            #puts "r #{r}"
            r = r.gsub(/[\s,]/ ,"")
            if r =~ /\A[-+]?[0-9]*\.?[0-9]+\Z/ 
              temp << r.to_f
            elsif r != 'NA' && r != '(1)' && r != '(2)'
              name << r
            elsif r != '(1)' && r != '(2)'
              temp << r
            end
              
            if i == row.length-1
              name = name.join(" ")
              temp.unshift(name).flatten
              row = temp
              second_temp << row
              temp = []
              name = []
            end
          end
          #puts "second temp #{second_temp}"
        end
        
          final_second = []
          head_second.each_with_index do |h, ind|
            second_temp.each_with_index do |f, i|
              if ind != 0 
                temp = [f[0], h, f[ind]]
                final_second << temp
              end
            end
          end
      
      final_first = final_first.push(*final_second)
      
      
      puts "Saving #{final_first.length} records....."
      
      ActiveRecord::Base.transaction do
        final_first.each do |f|
          if f[2] == 'NA'
            f[2] = 0
          end
            ActiveRecord::Base.connection.execute "INSERT INTO imf_data (date, currency_name, rate) VALUES ('#{f[1]}', '#{f[0]}', #{f[2]})"
        end
      end
      
      
#####UNCOMMENT FOR FULL DATA PULL
    rolling_date = rolling_date.next_month.end_of_month
    end #WHILE LOOP END
#####

  end #DEF END
end #WHOLE THING END

class String
    def numeric?
      Float(self) != nil rescue false
    end
end

