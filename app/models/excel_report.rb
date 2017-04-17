class ExcelReport < ApplicationRecord
  
  def self.update_analysis
    puts "Current time to track when it is running #{Time.now()}"
    package = Axlsx::Package.new
    workbook = package.workbook
    
    @alphabet = ('A'..'Z').to_a
    
    @rates = ImfDatum.order('currency_name, date').pluck('date, currency_name, rate')
    @monthly = ImfDatum
    .where("currency_name != ?",'')
    .group("currency_name, years, DATE_PART('month', date)")
    .order("currency_name, years, DATE_PART('month', date)")
    .pluck("currency_name, AVG(CASE WHEN rate > 0 THEN rate ELSE 0 END), DATE_PART('year', date) as years")
    
    @countries = ExportVol
    .from('export_vols a')
    .joins('left join country_mappings b ON b.country = a.country')
    .order('a.country')
    .pluck('distinct(a.country), b.currency')
    
    @mail_codes = ['C', 'E']
    
    if Date.today > Date.new(Date.today.year,4,1)
    @dates_back = Date.today.year - 2
    else
    @dates_back = Date.today.year - 3
    end
     
    @dates = ImfDatum
    .from('imf_data a')
    .joins("
        INNER JOIN country_mappings b ON LOWER(a.currency_name) = LOWER(b.currency)
        INNER JOIN export_vols c ON c.country = b.country AND to_char(to_date(c.month, 'Month'), 'MM')::int = date_part('month', a.date) AND c.year = date_part('year', a.date)
        ")
    .where("date_part('year', a.date) > ?", @dates_back)
    .order("date_part('year', a.date), date_part('month', a.date)")
    .pluck("DISTINCT(date_part('month', a.date)),c.month, date_part('year', a.date)")
    
    
    @hong_kong = OutboundGbsDatabase
    .from('outbound_gbs_databases a')
    .joins("inner join country_mappings b on b.country_code = a.dest
            inner join export_vols c ON c.country = b.country AND to_char(to_date(c.month, 'Month'), 'MM')::int = date_part('month', a.month) AND c.year = a.year")
    .where('dest = ? and mail_class_code IN (?) and c.year > ?', 'HK', @mail_codes, @dates_back)
    .group('dest, a.month, c.year, export_vol')
    .order("c.year, date_part('month', a.month)")
    .pluck("to_char(a.month, 'Month'), c.year, array_agg(distinct(pieces)), export_vol")
    
    @country_array = []
    @currency_array = []
      
      @countries.each do |cntry|
        @country_array << cntry[0]
        @currency_array << cntry[1]
      end
      
    
    #conditional format styling
    unprofitable = workbook.styles.add_style( :fg_color => "B22727", :bg_color => "FFD1D1", :type => :dxf )

    #WORKBOOK STYLING    
    workbook.styles do |s|
      @italics = s.add_style i: true, sz: 7
      @header = s.add_style alignment: {horizontal: :center}, b: true, sz: 10, bg_color: "C0C0C0"
      @merge = s.add_style alignment: {horizontal: :center}
      
      @date = s.add_style :num_fmt => 14, :border => { :style => :thin, :color =>"000000" }
      @decimal = s.add_style :num_fmt => 4, :border => { :style => :thin, :color =>"000000" }
      @number = s.add_style :num_fmt => 3, :border => { :style => :thin, :color =>"000000" }
      @percent = s.add_style :num_fmt => 10, :border => { :style => :thin, :color =>"000000" }
      
      @details = s.add_style sz: 7, b: true, alignment: {horizontal: :center, wrap_text: true}
      @title = s.add_style sz: 18, b: true, alignment: {horizontal: :center}
      
      @row_headers = s.add_style sz: 10, b: true, alignment: {horizontal: :left, wrap_text: true}, :border => { :style => :thin, :color =>"000000" } 
      @gridstyle_border =  s.add_style :border => { :style => :thin, :color =>"000000" }
      @colorless = s.add_style :fg_color => "FFFFFF"
      @cell_rotated_text_style = s.add_style sz: 8, :alignment => {:textRotation => 90}, :border => { :style => :thin, :color =>"000000" }
      @shrink_wrap = s.add_style sz: 8, alignment: {wrap_text: true, horizontal: :center}, :border => { :style => :thin, :color =>"000000" }
      @wrap_text_border = s.add_style alignment: {wrap_text: true}, :border => { :style => :thin, :color =>"000000" }
      @no_width = s.add_style :width => 1
      
      
    #IMF ORIGINAL DATA SHEET
    workbook.add_worksheet name: "IMF Data" do |sheet|
      #sheet.add_row
      
      sheet.add_row ["Date", "Currency Name", "Rate"], :style => @header
      @rates.each do |st|
        if st[2] == 0
          st[2] = 'NA'
        end
        sheet.add_row [st[0], st[1], st[2]], :style => [@date, @gridstyle_border, @decimal]
        sheet.column_widths 10, 20, 10
      end
      
      
      #sheet.row_style 0, @header, col_offset: 1
      #sheet.col_style 0, @italics, row_offset: 1
    end
    
    #IMF MONTHLY DATA SHEET
    workbook.add_worksheet name: "IMF Monthly Data" do |w|
     w.add_row ["Year", "Currency", "Helper", "Month"], :style => @header
     w.merge_cells("D1:O1")
     #w.col_style 2, @merge, row_offset: 0
     w.add_row ["","", "", "January","February","March","April","May","June","July","August","September","October","November","December"], style: @header
     
     current_year = ''
     final = []
     temp = []
     @monthly.each_with_index do |m, i|
      
      if current_year == ''
        current_year = 1995
        temp = [m[0]]
      end
      
      
      if m[1] == 0
        m[1] = 'NA'
      end
      
      if i == @monthly.length-1
        temp << [m[1], current_year]
        final << temp.flatten
      elsif current_year != m[2]
          temp << current_year
          current_year = m[2]
          final << temp.flatten
          temp = [m[0], m[1]]
      else 
        temp << m[1]
      end
      
     end
        
    
     final.each do |f|
       
       temp_arr = []
       temp_year = 0
       temp_name = ''
       temp_helper = ''
       
       f.each do |element|
         if element.is_a?(Numeric) && element == f[-1]
           temp_year = element.to_i.to_s
         elsif element.is_a?(String) && element.length > 2
           temp_name = element
         else
           temp_arr << element
         end
       
       end
       
       temp_helper = temp_name+"_"+temp_year
       temp_arr.unshift([temp_year, temp_name, temp_helper])
       
       w.add_row [].push(temp_arr).flatten
       
     end
    end #end of worksheet IMF monthly
    
    #IMF MONTHLY CONVERSIONS
    workbook.add_worksheet name: "IMF Monthly Conversions" do |con|
      conversions = ["Euro", "U.K. Pound Sterling", "Australian Dollar", "New Zealand Dollar"]
      
      con.add_row ["Currency", "Date", "Monthly Avg.", "IMF Conversion"], :style => @header
      conversions.each do |cur|
        @dates.each do |d|
          con.add_row [cur, d[1][0..2]+"-"+d[2].to_s[-4..-3], "=VLOOKUP(\"#{cur}\"&\"_\"&\"#{d[2].to_i}\",'IMF Monthly Data'!$C$3:$O$99999, MATCH(\"#{d[1]}\",'IMF Monthly Data'!$C$2:$O$2,0),FALSE)", "=1/VLOOKUP(\"#{cur}\"&\"_\"&\"#{d[2].to_i}\",'IMF Monthly Data'!$C$3:$O$99999, MATCH(\"#{d[1]}\",'IMF Monthly Data'!$C$2:$O$2,0),FALSE)"], :style => @gridstyle_border
        end  
      end
      
      con.col_style 2, @decimal, row_offset: 1
      con.col_style 3, @decimal, row_offset: 1
    end
    
    #IMF 5 Month LOOKBACK
     workbook.add_worksheet name: "Lookback 5 Months" do |five_page|
     @five_months = ImfDatum
     .limit(6)
     .order("(date_part('year', date)) desc, date_part('month', date) desc")
     .pluck("DISTINCT(to_char(date, 'Mon')), date_part('year', date), to_char(date, 'Month'),date_part('month', date)")

     @five_months = @five_months.reverse
     @five_months.pop
     
     five_page.add_row [""]
     
     five_page.add_data_validation("E2", {
      :type => :list,
      :formula1 => 'C11:M11',
      :showDropDown => false,
      :showInputMessage => true,
      :promptTitle => 'Year SPLY',
      :prompt => 'Choose currency to view (Default: Canadian Dollar)'
      })
     
     five_page.add_row [" ", "Year SPLY by month", "", "Currency:", "Canadian Dollar", ""], :style => [nil, @header, nil, nil, @gridstyle_border, @gridstyle_border]  
     five_page.add_row [" ","Date", "Variance"], :style => [nil, @row_headers, @row_headers]
     five_page.merge_cells "E2:F2"
     
     @five_months.each do |o|
        @text_year = o[1].to_i.to_s
        @date = o[0]+"-"+@text_year[-2..-1]
        five_page.add_row [" ", @date, "=IF($E$2=\"\",VLOOKUP(\"#{@date}\",$B$11:$M$16,MATCH($C$11,$B$11:$M$11,0),FALSE),VLOOKUP(\"#{@date}\",$B$11:$M$16,MATCH($E$2,$B$11:$M$11,0),FALSE))"], :style => [nil, @gridstyle_border, @percent]
      end
     
     five_page.add_row []
     five_page.add_row []
     five_page.add_row [" ","Date", "Canadian Dollar","Australian Dollar","U.K. Pound Sterling","Chinese Yuan","Japanese Yen","Euro","Russian Ruble","Korean Won","Brazilian Real","Mexican Peso","Hong Kong Dollar"], :style => @colorless
     
     @five_months.each_with_index do |d, ind|
       @text_year = d[1].to_i.to_s  
       
       if d[3] < 5
         @months_back_year = (d[1]-1).to_i.to_s
       else
         @months_back_year = @text_year
       end
       
       @strip_month = d[2].gsub(/\s+/, "")
       @current_date = Date.new(d[1],d[3],1)
       @strip_5_back = @current_date.months_ago(5)
       @months_back = @strip_5_back.strftime("%B")
       
       #incrementing the alphabet, starting with row C 
       @alphabet_count = 2
       temp_array = []
       
       #number comes from count of currencies used minus hong kong
       10.times do 
         temp_array << "=(VLOOKUP(#{@alphabet[@alphabet_count]}11&\"_\"&\"#{@text_year}\",'IMF Monthly Data'!$C$2:$O$99999,MATCH(\"#{@strip_month}\",'IMF Monthly Data'!$C$2:$O$2,0),FALSE)-VLOOKUP(C11&\"_\"&\"#{@text_year}\",'IMF Monthly Data'!$C$2:$O$99999,MATCH(\"#{@months_back}\",'IMF Monthly Data'!$C$2:$O$2,0),FALSE))/VLOOKUP(C11&\"_\"&\"#{@months_back_year}\",'IMF Monthly Data'!$C$2:$O$99999,MATCH(\"#{@months_back}\",'IMF Monthly Data'!$C$2:$O$2,0),FALSE)"
         
         @alphabet_count += 1
       end
       
       five_page.add_row [" ", d[0]+"-"+@text_year[-2..-1]].push(temp_array).flatten, :style => @colorless

      end
        
      chart = five_page.add_chart(Axlsx::LineChart, :start_at=> "E4", :end_at=> "M28", :show_legend => false, :title=>"Monthly Variance in Currency vs. USD")
            chart.add_series :data => five_page["C4:C8"], :labels => five_page["B4:B8"], :show_marker => true
            chart.valAxis.gridlines = false
            chart.catAxis.gridlines = false
      
     
     
    end
     
    #IMF ONE YEAR LOOKBACK
     workbook.add_worksheet name: "Lookback 1 Year" do |l|
     @one_year = ImfDatum
     .limit(13)
     .order("(date_part('year', date)) desc, date_part('month', date) desc")
     .pluck("DISTINCT(to_char(date, 'Mon')), date_part('year', date), to_char(date, 'Month'),date_part('month', date)")

     @one_year = @one_year.reverse
     @one_year.pop
     
     l.add_row [""]
     
      l.add_data_validation("E2", {
      :type => :list,
      :formula1 => 'C18:M18',
      :showDropDown => false,
      :showInputMessage => true,
      :promptTitle => 'Year SPLY',
      :prompt => 'Choose currency to view (Default: Canadian Dollar)'
      })
      
     l.add_row [" ", "Year SPLY by month", "", "Currency:", "Canadian Dollar", ""], :style => [nil, @header, nil, nil, @gridstyle_border, @gridstyle_border]  
     l.add_row [" ","Date", "Variance"], :style => [nil, @row_headers, @row_headers]
     l.merge_cells "E2:F2"
     
     @one_year.each do |o|
        @text_year = o[1].to_i.to_s
        @date = o[0]+"-"+@text_year[-2..-1]
        l.add_row [" ", @date, "=IF($E$2=\"\",VLOOKUP(\"#{@date}\",$B$18:$M$30,MATCH($C$18,$B$18:$M$18,0),FALSE),VLOOKUP(\"#{@date}\",$B$18:$M$30,MATCH($E$2,$B$18:$M$18,0),FALSE))"], :style => [nil, @gridstyle_border, @percent]
      end
     
     l.add_row []
     l.add_row []
     l.add_row [" ","Date", "Canadian Dollar","Australian Dollar","U.K. Pound Sterling","Chinese Yuan","Japanese Yen","Euro","Russian Ruble","Korean Won","Brazilian Real","Mexican Peso","Hong Kong Dollar"], :style => @colorless
     
     @one_year.each_with_index do |d, ind|
       #puts "ORIG #{d}"
       @text_year = d[1].to_i.to_s
       @strip_month = d[2].gsub(/\s+/, "")
       #puts "strip month #{@strip_month}"
       
       
       #incrementing the alphabet, starting with row C 
       @alphabet_count = 2
       temp_array = []
       
       #number comes from count of currencies used minus hong kong
       10.times do 
         temp_array << "=(VLOOKUP(#{@alphabet[@alphabet_count]}18&\"_\"&\"#{(d[1]).to_i}\",'IMF Monthly Data'!$C$2:$O$99999,MATCH(\"#{@strip_month}\",'IMF Monthly Data'!$C$2:$O$2,0),FALSE)-VLOOKUP(C18&\"_\"&\"#{(d[1]-1).to_i}\",'IMF Monthly Data'!$C$2:$O$99999,MATCH(\"#{@strip_month}\",'IMF Monthly Data'!$C$2:$O$2,0),FALSE))/VLOOKUP(C18&\"_\"&\"#{(d[1]-2).to_i}\",'IMF Monthly Data'!$C$2:$O$99999,MATCH(\"#{@strip_month}\",'IMF Monthly Data'!$C$2:$O$2,0),FALSE)"
         @alphabet_count += 1
       end
       
       l.add_row [" ", d[0]+"-"+@text_year[-2..-1]].push(temp_array).flatten, :style => @colorless
       
      end
     
      
      
     
      
      
      
      chart = l.add_chart(Axlsx::LineChart, :start_at=> "E4", :end_at=> "M28", :show_legend => false, :title=>"Monthly Variance in Currency vs. USD")
            chart.add_series :data => l["C4:C15"], :labels => l["B4:B15"], :show_marker => true
            chart.valAxis.gridlines = false
            chart.catAxis.gridlines = false
      
     
     
    end
     
    #IMF TWO YEAR LOOKBACK
     workbook.add_worksheet name: "Lookback 2 Year" do |two|
     @two_year = ImfDatum
     .limit(25)
     .order("(date_part('year', date)) desc, date_part('month', date) desc")
     .pluck("DISTINCT(to_char(date, 'Mon')), date_part('year', date), to_char(date, 'Month'),date_part('month', date)")

     @two_year = @two_year.reverse
     @two_year.pop
     
     two.add_row [""]
     
      two.add_data_validation("E2", {
      :type => :list,
      :formula1 => 'C30:M30',
      :showDropDown => false,
      :showInputMessage => true,
      :promptTitle => 'Year SPLY',
      :prompt => 'Choose currency to view (Default: Canadian Dollar)'
      })
      
     two.add_row [" ", "Year SPLY by month", "", "Currency:", "Canadian Dollar", ""], :style => [nil, @header, nil, nil, @gridstyle_border, @gridstyle_border]  
     two.add_row [" ","Date", "Variance"], :style => [nil, @row_headers, @row_headers]
     two.merge_cells "E2:F2"
     
     @two_year.each do |o|
        @text_year = o[1].to_i.to_s
        @date = o[0]+"-"+@text_year[-2..-1]
        two.add_row [" ", @date, "=IF($E$2=\"\",VLOOKUP(\"#{@date}\",$B$30:$M$54,MATCH($C$30,$B$30:$M$30,0),FALSE),VLOOKUP(\"#{@date}\",$B$30:$M$54,MATCH($E$2,$B$30:$M$30,0),FALSE))"], :style => [nil, @gridstyle_border, @percent]
      end
     
     two.add_row []
     two.add_row []
     two.add_row [" ","Date", "Canadian Dollar","Australian Dollar","U.K. Pound Sterling","Chinese Yuan","Japanese Yen","Euro","Russian Ruble","Korean Won","Brazilian Real","Mexican Peso","Hong Kong Dollar"], :style => @colorless
     
     @two_year.each_with_index do |d, ind|
       #puts "ORIG #{d}"
       @text_year = d[1].to_i.to_s
       @strip_month = d[2].gsub(/\s+/, "")
       #puts "strip month #{@strip_month}"
       
       #incrementing the alphabet, starting with row C 
       @alphabet_count = 2
       temp_array = []
       
       #number comes from count of currencies used minus hong kong
       10.times do 
         temp_array << "=(VLOOKUP(#{@alphabet[@alphabet_count]}30&\"_\"&\"#{(d[1]).to_i}\",'IMF Monthly Data'!$C$2:$O$99999,MATCH(\"#{@strip_month}\",'IMF Monthly Data'!$C$2:$O$2,0),FALSE)-VLOOKUP(C30&\"_\"&\"#{(d[1]-2).to_i}\",'IMF Monthly Data'!$C$2:$O$99999,MATCH(\"#{@strip_month}\",'IMF Monthly Data'!$C$2:$O$2,0),FALSE))/VLOOKUP(C30&\"_\"&\"#{(d[1]-2).to_i}\",'IMF Monthly Data'!$C$2:$O$99999,MATCH(\"#{@strip_month}\",'IMF Monthly Data'!$C$2:$O$2,0),FALSE)"
         @alphabet_count += 1
       end
       
       two.add_row [" ", d[0]+"-"+@text_year[-2..-1]].push(temp_array).flatten, :style => @colorless
                  
      end   
      
      
      chart = two.add_chart(Axlsx::LineChart, :start_at=> "E4", :end_at=> "M28", :show_legend => false, :title=>"Monthly Variance in Currency vs. USD")
            chart.add_series :data => two["C4:C27"], :labels => two["B4:B27"], :show_marker => true
            chart.valAxis.gridlines = false
            chart.catAxis.gridlines = false
            chart.catAxis.label_rotation = -90
      
     
     
    end 
     
    #Individiual Country
    workbook.add_worksheet(name: "Country") do |page|
      
      ###Hide gridlines on page
      page.sheet_view.show_grid_lines= false
    
      page.merge_cells "B3:C3"
      
      page.add_data_validation("B3", {
      :type => :list,
      :formula1 => "A4:X4",
      :showDropDown => false,
      :showInputMessage => true,
      :promptTitle => 'Country List',
      :prompt => 'Please select a Country!'
      })
      
      page.add_row ["Country Exchange Rate"], :style => @title
      page.add_row [" "]
      page.add_row ["Country:", "Australia"]
      page.add_row [].push(@country_array).flatten, :style => @colorless
      page.add_row ["Date", "Exchange \x0D Rate", "EMS", "% Change in EMS \x0DVolumes", "Parcels", "% Change in Parcel \x0DVolumes", "Postal Service \x0DParcel Volume", "% Change in volume from \x0Dmonth to month", "US Export \x0DVolume (in millions)"], style: @row_headers, :height => 60, :widths => [5, 5, 5, 15, 5, 15, 5, 15, 5]
      
      
      @dates.each_with_index do |d, i|
        @alphabet_count = 1
        temp_array = []
        
        8.times do
        temp_array << "=VLOOKUP($B$3&\"_\"&$A#{i+6},'Country Raw Data'!$C$1:$L$99999,MATCH(#{@alphabet[@alphabet_count]}$5,'Country Raw Data'!$C$1:$L$1,0),FALSE)"
        @alphabet_count += 1  
        end
        
        page.add_row [d[1][0..2]+"-"+d[2].to_s[-4..-3]].push(temp_array).flatten
        
      end
      
      #Column Style
      page.column_widths 10, 10, 10, 10, 10, 10, 10, 10, 11, 10
      page.col_style 1, @decimal, row_offset: 5
      page.col_style 2, @number, row_offset: 5
      page.col_style 3, @percent, row_offset: 5
      page.col_style 4, @number, row_offset: 5
      page.col_style 5, @percent, row_offset: 5
      page.col_style 6, @number, row_offset: 5
      page.col_style 7, @percent, row_offset: 5
      page.col_style 8, @decimal, row_offset: 5
         
      #Conditional Formatting
          page.add_conditional_formatting "A6:I#{@dates.length+5}", { :type => :cellIs,
                                             :operator => :lessThan,
                                             :formula => '0',
                                             :dxfId => unprofitable,
                                             :priority => 1 }
          
      #CHARTS ADJUSTED FOR NUMBER OF ROW ENTRIES
          #if  != 'Hong Kong'
            
            chart = page.add_chart(Axlsx::LineChart, :start_at=> "K1", :end_at=> "S12", :show_legend => false, :title=>"Exchange Rate Trends")
            chart.add_series :data => page["B6:B#{@dates.length+5}"], :labels => page["A6:A#{@dates.length+5}"], :title => page["B5"], :show_marker => true
            chart.catAxis.label_rotation = -90
            chart.valAxis.gridlines = false
            chart.catAxis.gridlines = false
            
            page.add_chart(Axlsx::Bar3DChart, :start_at => "K13", :end_at => "S30", :title=> "Outbound Pieces", :barDir => :col) do |bars|
              bars.add_series :data => page["C6:C#{@dates.length+5}"], :labels => page["A6:A#{@dates.length+5}"], :title => page["C5"]
              bars.add_series :data => page["E6:E#{@dates.length+5}"], :labels => page["A6:A#{@dates.length+5}"], :title => page["E5"]
              bars.catAxis.label_rotation = -90
              bars.valAxis.gridlines = false
              bars.catAxis.gridlines = false
            end
            
            page.add_chart(Axlsx::Bar3DChart, :start_at => "K30", :end_at => "S41", :title=> "Exports (in millions)", :barDir => :col, :show_legend => false) do |exports|
              exports.add_series :data => page["I6:I#{@dates.length+5}"], :labels => page["A6:A#{@dates.length+5}"], :title => page["I5"], :color => "FF0000"
              exports.catAxis.label_rotation = -90
              exports.valAxis.gridlines = false
              exports.catAxis.gridlines = false
            end
          
          dual_line = page.add_chart(Axlsx::LineChart, :start_at=> "S1", :end_at=> "AE25", :title=>"Export Trends")
          dual_line.add_series :data => page["G6:G#{@dates.length+5}"], :labels => page["A6:A#{@dates.length+5}"], :title => page["G5"], :show_marker => true
          dual_line.add_series :data => page["I6:I#{@dates.length+5}"], :labels => page["A6:A#{@dates.length+5}"], :title => page["I5"], :show_marker => true, :on_primary_axis => false  
          dual_line.catAxis.label_rotation = -90
          dual_line.valAxis.gridlines = false
          dual_line.catAxis.gridlines = false 
    end
    
    #CORRELATION PAGE
    workbook.add_worksheet name: "Correlation" do |c|
      
      c.sheet_view.show_grid_lines= false
      
      
      #Ending point of array/num of rows needed plus offset of titles/headers
      #FOR USE ON CORRELATION PAGE ONLY
      @end_count = @dates.length + 4
      
      c.add_row [" ", "Country Exchange Rates and Correlation"], :style =>[nil, @title]
      c.add_row []
      c.add_row ["", "Date"].push(@country_array).flatten, :style => @wrap_text_border
      c.add_row [" ", " "].push(@currency_array).flatten, :style => @wrap_text_border
      
      counter = 6
      
      #+6 because country tab starts on 6 and then +1 to offset index
      while counter < @dates.length + 6
        
        @alphabet_count = 2
        @temp_array = []
        24.times do
          @temp_array << "=VLOOKUP(#{@alphabet[@alphabet_count]}$3&\"_\"&Correlation!$B#{counter-1},'Country Raw Data'!$C$1:$L$99999,3,FALSE)"
          @alphabet_count += 1
        end
        
        c.add_row [" ", "='Country'!A#{counter}"].push(@temp_array).flatten
        counter += 1
      end
      
      @alphabet_count = 2
      @total_array = []
      @ems_array = []
      @parcel_array = []
      @export_array = []
      @exchange_array = []
      @difference_array = []
      @sample_array = []
      @population_array = []
      @max_array = []
      @min_array = []
      @avg_array = []
      @stdev_array = []
      
      24.times do
        @total_array << "=IFERROR(CORREL(INDIRECT(\"'Country Raw Data'!\"&\"E\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B5,'Country Raw Data'!$C$1:$C$99999,0))):INDIRECT(\"'Country Raw Data'!\"&\"E\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B#{@end_count},'Country Raw Data'!$C$1:$C$99999,0))), INDIRECT(\"'Country Raw Data'!\"&\"J\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B5,'Country Raw Data'!$C$1:$C$99999,0))):INDIRECT(\"'Country Raw Data'!\"&\"J\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B#{@end_count},'Country Raw Data'!$C$1:$C$99999,0)))),0)"
        @ems_array <<  "=IFERROR(CORREL(INDIRECT(\"'Country Raw Data'!\"&\"E\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B5,'Country Raw Data'!$C$1:$C$99999,0))):INDIRECT(\"'Country Raw Data'!\"&\"E\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B#{@end_count},'Country Raw Data'!$C$1:$C$99999,0))), INDIRECT(\"'Country Raw Data'!\"&\"F\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B5,'Country Raw Data'!$C$1:$C$99999,0))):INDIRECT(\"'Country Raw Data'!\"&\"F\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B#{@end_count},'Country Raw Data'!$C$1:$C$99999,0)))),0)"
        @parcel_array << "=IFERROR(CORREL(INDIRECT(\"'Country Raw Data'!\"&\"E\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B5,'Country Raw Data'!$C$1:$C$99999,0))):INDIRECT(\"'Country Raw Data'!\"&\"E\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B#{@end_count},'Country Raw Data'!$C$1:$C$99999,0))), INDIRECT(\"'Country Raw Data'!\"&\"H\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B5,'Country Raw Data'!$C$1:$C$99999,0))):INDIRECT(\"'Country Raw Data'!\"&\"H\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B#{@end_count},'Country Raw Data'!$C$1:$C$99999,0)))),0)"
        @export_array << "=IFERROR(CORREL(INDIRECT(\"'Country Raw Data'!\"&\"J\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B5,'Country Raw Data'!$C$1:$C$99999,0))):INDIRECT(\"'Country Raw Data'!\"&\"J\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B#{@end_count},'Country Raw Data'!$C$1:$C$99999,0))), INDIRECT(\"'Country Raw Data'!\"&\"L\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B5,'Country Raw Data'!$C$1:$C$99999,0))):INDIRECT(\"'Country Raw Data'!\"&\"L\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B#{@end_count},'Country Raw Data'!$C$1:$C$99999,0)))),0)"
        @exchange_array << "=IFERROR(CORREL(INDIRECT(\"'Country Raw Data'!\"&\"E\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B5,'Country Raw Data'!$C$1:$C$99999,0))):INDIRECT(\"'Country Raw Data'!\"&\"E\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B#{@end_count},'Country Raw Data'!$C$1:$C$99999,0))), INDIRECT(\"'Country Raw Data'!\"&\"L\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B5,'Country Raw Data'!$C$1:$C$99999,0))):INDIRECT(\"'Country Raw Data'!\"&\"L\"&(MATCH(#{@alphabet[@alphabet_count]}3&\"_\"&B#{@end_count},'Country Raw Data'!$C$1:$C$99999,0)))),0)"
        @difference_array << "=#{@alphabet[@alphabet_count]}#{@end_count+4}-#{@alphabet[@alphabet_count]}#{@end_count+5}"
        @sample_array << "=VAR(#{@alphabet[@alphabet_count]}5:#{@alphabet[@alphabet_count]}#{@end_count})"
        @population_array << "=VARP(#{@alphabet[@alphabet_count]}5:#{@alphabet[@alphabet_count]}#{@end_count})"
        @max_array << "=MAX(#{@alphabet[@alphabet_count]}5:#{@alphabet[@alphabet_count]}#{@end_count})"
        @min_array << "=MIN(#{@alphabet[@alphabet_count]}5:#{@alphabet[@alphabet_count]}#{@end_count})"
        @avg_array << "=AVERAGE(#{@alphabet[@alphabet_count]}5:#{@alphabet[@alphabet_count]}#{@end_count})"
        @stdev_array << "=STDEV(#{@alphabet[@alphabet_count]}5:#{@alphabet[@alphabet_count]}#{@end_count})"
        @alphabet_count += 1
      end
      
      c.add_row ["Correlation", "Total Volume"].push(@total_array).flatten, :style => [@cell_rotated_text_style]
      c.add_row [" ", "EMS"].push(@ems_array).flatten
      c.add_row [" ", "Parcels"].push(@parcel_array).flatten
      c.add_row ["Correlation of U.S.\x0D\x0AExports to PS Vol", ""].push(@export_array).flatten, :style => [@shrink_wrap], :height => 25      
      c.add_row ["Correlation of Exchange Rates \x0D\x0A to US Imports", ""].push(@exchange_array).flatten, :style => [@shrink_wrap], :height => 25       
      c.add_row ["Difference between correlation from exchange rates PS Vol and exchange rates to US Exports", ""].push(@difference_array).flatten, :style => [@shrink_wrap], :height => 25
      c.add_row []
      c.add_row []
      c.add_row [" ", "Variance as a sample"].push(@sample_array).flatten, :style => [nil, @wrap_text_border]
      c.add_row [" ", "Variance as a population"].push(@population_array).flatten, :style => [nil, @wrap_text_border]
      c.add_row [" ", "Maximum"].push(@max_array).flatten
      c.add_row [" ", "Minimum"].push(@min_array).flatten
      c.add_row [" ", "Average"].push(@avg_array).flatten
      c.add_row [" ", "Standard Deviation"].push(@stdev_array).flatten
      c.merge_cells "A#{@end_count+1}:A#{@end_count+3}"
      c.merge_cells "A#{@end_count+4}:B#{@end_count+4}"
      c.merge_cells "A#{@end_count+5}:B#{@end_count+5}"
      c.merge_cells "A#{@end_count+6}:B#{@end_count+6}"
      c.merge_cells "B1:I1"
                
      #Column Styles
          c.column_widths 3, 21, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10 ,10 ,10 ,10 ,10 ,10 ,10 ,10 ,10 ,10
          c.col_style 1, @decimal, row_offset: 4
          c.col_style 2, @decimal, row_offset: 4
          c.col_style 3, @decimal, row_offset: 4
          c.col_style 4, @decimal, row_offset: 4
          c.col_style 5, @decimal, row_offset: 4
          c.col_style 6, @decimal, row_offset: 4
          c.col_style 7, @decimal, row_offset: 4
          c.col_style 8, @decimal, row_offset: 4
          c.col_style 9, @decimal, row_offset: 4
          c.col_style 10, @decimal, row_offset: 4
          c.col_style 11, @decimal, row_offset: 4
          c.col_style 12, @decimal, row_offset: 4
          c.col_style 13, @decimal, row_offset: 4
          c.col_style 14, @decimal, row_offset: 4
          c.col_style 15, @decimal, row_offset: 4
          c.col_style 16, @decimal, row_offset: 4
          c.col_style 17, @decimal, row_offset: 4
          c.col_style 18, @decimal, row_offset: 4
          c.col_style 19, @decimal, row_offset: 4
          c.col_style 20, @decimal, row_offset: 4
          c.col_style 21, @decimal, row_offset: 4
          c.col_style 22, @decimal, row_offset: 4
          c.col_style 23, @decimal, row_offset: 4
          c.col_style 24, @decimal, row_offset: 4
          c.col_style 25, @decimal, row_offset: 4
          c.col_style 26, @decimal, row_offset: 4
      
      #Conditional Formatting
        c.add_conditional_formatting "C#{@end_count+1}:Z#{@end_count+3}", { :type => :cellIs,
                                           :operator => :lessThan,
                                           :formula => '-0.5',
                                           :dxfId => unprofitable,
                                           :priority => 1 }
        
        c.add_conditional_formatting "C#{@end_count+2}:Z#{@end_count+3}", { :type => :cellIs,
                                           :operator => :lessThan,
                                           :formula => '0',
                                           :dxfId => unprofitable,
                                           :priority => 1 }
          
          
    end
    
    ##Volume Changes
    workbook.add_worksheet name: "Volume Changes" do |tab|
      
      
      tab.sheet_view.show_grid_lines= false
      
      #3 to offset headers and whatnot
      @length_count = @dates.length
      
      tab.add_row ["", "Country Volume Percentage Changes from Month to Month"], :style => @title
      tab.add_row []
      tab.add_row [" "].push(@country_array).flatten, :style => @row_headers
      
      @position = 4
      @dates.each do |data|
        
        @alphabet_count = 1
        @temp_array = []
        
        24.times do
          @temp_array << "=VLOOKUP($#{@alphabet[@alphabet_count]}3&\"_\"&A#{@position},'Country Raw Data'!$C$1:$L$99999,9,FALSE)"
          @alphabet_count += 1
        end
        
        
        tab.add_row [data[1][0..2]+"-"+data[2].to_s[-4..-3]].push(@temp_array).flatten, :style => @percent
        
        @position += 1
      end
    
    
    tab.add_row []
    tab.add_row ["", "Country EMS Volume Percentage Changes from Month to Month"], :style => @title
    tab.add_row []
    tab.add_row [" "].push(@country_array).flatten, :style => @row_headers
    
      @position = 4
    
      @dates.each do |data|
        
        @alphabet_count = 1
        @temp_array = []
        24.times do
          @temp_array << "=VLOOKUP(#{@alphabet[@alphabet_count]}$3&\"_\"&A#{@position},'Country Raw Data'!$C$1:$L$99999,5,FALSE)"
          @alphabet_count += 1
        end
        
        tab.add_row [data[1][0..2]+"-"+data[2].to_s[-4..-3]].push(@temp_array).flatten, :style => @percent

        @position += 1
      end
      
      
    tab.add_row []
    tab.add_row ["", "Country Parcel Volume Percentage Changes from Month to Month"], :style => @title
    tab.add_row []
    tab.add_row [" "].push(@country_array).flatten, :style => @row_headers
     
      @position = 4
    
      @dates.each do |data|
        
        @alphabet_count = 1
        @temp_array = []
        
        24.times do
          @temp_array << "=VLOOKUP(#{@alphabet[@alphabet_count]}$3&\"_\"&A#{@position},'Country Raw Data'!$C$1:$L$99999,7,FALSE)"
          @alphabet_count += 1
        end
        
        tab.add_row [data[1][0..2]+"-"+data[2].to_s[-4..-3]].push(@temp_array).flatten, :style => @percent
                     
        
        @position += 1
      end
          
        tab.merge_cells "B1:Y1"
        tab.merge_cells "B#{@length_count+5}:Y#{@length_count+5}"
        tab.merge_cells "B#{(@length_count*2)+9}:Y#{(@length_count*2)+9}"
        tab.column_widths 10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10,10
          
        #Conditional Formatting
        tab.add_conditional_formatting "A4:Y#{@length_count+3}", { :type => :cellIs,
                                           :operator => :lessThan,
                                           :formula => '0',
                                           :dxfId => unprofitable,
                                           :priority => 1 }
        
        tab.add_conditional_formatting "A#{@length_count+3+5}:Y#{(@length_count+3+4)+@length_count}", { :type => :cellIs,
                                           :operator => :lessThan,
                                           :formula => '0',
                                           :dxfId => unprofitable,
                                           :priority => 1 }
      
        tab.add_conditional_formatting "A#{(@length_count*2+3+4)+5}:Y#{(@length_count*2+3+4)+5+@length_count}", { :type => :cellIs,
                                           :operator => :lessThan,
                                           :formula => '0',
                                           :dxfId => unprofitable,
                                           :priority => 1 }
        
    end
    
    #COUNTRY RAW DATA, LEADS TO COUNTRY DATA
    workbook.add_worksheet name: 'Country Raw Data', state: :hidden do |sheet|
      sheet.add_row ["Country", "Currency", "Helper", "Date", "Exchange \x0D Rate", "EMS", "% Change in EMS \x0DVolumes", "Parcels", "% Change in Parcel \x0DVolumes", "Postal Service \x0DParcel Volume", "% Change in volume from \x0Dmonth to month", "US Export \x0DVolume (in millions)"], style: @row_headers, :height => 60, :widths => [5, 5, 5, 15, 5, 15, 5, 15, 5]
      @countries.each_with_index do |country, i|
        #puts "CURRENT COUNTRY #{country[0]}"
        
        
        if country[0] != 'Hong Kong'
        @country_data = ImfDatum
          .from('imf_data a')
          .joins("
              INNER JOIN country_mappings b ON LOWER(a.currency_name) = LOWER(b.currency)
              INNER JOIN export_vols c ON c.country = b.country AND to_char(to_date(c.month, 'Month'), 'MM')::int = date_part('month', a.date) AND c.year = date_part('year', a.date)
              INNER JOIN outbound_gbs_databases d ON d.dest = b.country_code AND date_part('month', d.month) = date_part('month', a.date) AND d.year = date_part('year', a.date)
              ")
          .where("mail_class_code IN (?) AND rate != ? and date_part('year', a.date) > ? AND c.country = ?", @mail_codes, 0, @dates_back, country[0])
          .group("c.country, currency_name, currency, date_part('month', a.date), date_part('year', a.date), export_vol, c.month")
          .order("c.country, currency_name, date_part('year', a.date), date_part('month', a.date)")
          .pluck("c.country, currency_name, AVG(rate), c.month, date_part('year', a.date), export_vol, array_agg(DISTINCT(pieces))")
        end
        
        if @counter.nil?
            @counter = 1  
        else
            @counter = (@country_data.length * i)+1
        end
        
        if country[0] == 'Hong Kong'
              @hong_kong.each_with_index do |h, ind_count|
                h = h.flatten
                if ind_count == 0
                sheet.add_row ["Hong Kong", "Hong Kong Dollar", "=\"Hong Kong\"&\"_\"&\"#{h[0][0..2]}\"&\"-\"&\"#{h[1].to_s[-2..-1]}\"", h[0][0..2] +"-"+ h[1].to_s[-2..-1], 7.80, h[2], "", h[3], "", h[2]+h[3], "", h[4]], :style => @gridstyle_border
                else
                sheet.add_row ["Hong Kong", "Hong Kong Dollar", "=\"Hong Kong\"&\"_\"&\"#{h[0][0..2]}\"&\"-\"&\"#{h[1].to_s[-2..-1]}\"", h[0][0..2]+"-"+h[1].to_s[-2..-1], 7.80, h[2], "=(#{h[2]}-F#{@counter})/F#{@counter}", h[3], "=(#{h[3]}-H#{@counter})/H#{@counter}" ,h[2]+h[3], "=(#{h[2]+h[3]}-J#{@counter})/J#{@counter}",h[4]], :style => @gridstyle_border    
                end
                @counter += 1
              end  
        else
          #PULL DATA FOR CURRENT COUNTRY  
          
        
          @country_data.each_with_index do |c, index|
            c = c.flatten
            #puts "c[0] #{c[0]} country[0] #{country[0]} index #{index} COUNTER #{@counter}"
              if c[1] == 'Euro' || c[1] == 'Australian Dollar' || c[1] == 'U.K. Pound Sterling' || c[1] == 'New Zealand Dollar'
                if @counter == (@country_data.length * i)+1
                sheet.add_row [c[0], c[1], "=\"#{c[0]}\"&\"_\"&\"#{c[3][0..2]}\"&\"-\"&\"#{c[4].to_s[-4..-3]}\"", c[3][0..2]+"-"+c[4].to_s[-4..-3], 1/c[2], c[6], "", c[7], "",c[6]+c[7], "",c[5]], :style => @gridstyle_border
                else
                sheet.add_row [c[0], c[1], "=\"#{c[0]}\"&\"_\"&\"#{c[3][0..2]}\"&\"-\"&\"#{c[4].to_s[-4..-3]}\"", c[3][0..2]+"-"+c[4].to_s[-4..-3], 1/c[2], c[6], "=(#{c[6]}-F#{@counter})/F#{@counter}", c[7], "=(#{c[7]}-H#{@counter})/H#{@counter}" ,c[6]+c[7], "=(#{c[6]+c[7]}-J#{@counter})/J#{@counter}",c[5]], :style => @gridstyle_border
                end
              else
                if @counter == (@country_data.length * i)+1
                sheet.add_row [c[0], c[1], "=\"#{c[0]}\"&\"_\"&\"#{c[3][0..2]}\"&\"-\"&\"#{c[4].to_s[-4..-3]}\"", c[3][0..2]+"-"+c[4].to_s[-4..-3], c[2], c[6], "", c[7], "",c[6]+c[7], "",c[5]], :style => @gridstyle_border
                else
                sheet.add_row [c[0], c[1], "=\"#{c[0]}\"&\"_\"&\"#{c[3][0..2]}\"&\"-\"&\"#{c[4].to_s[-4..-3]}\"", c[3][0..2]+"-"+c[4].to_s[-4..-3], c[2], c[6], "=(#{c[6]}-F#{@counter})/F#{@counter}", c[7], "=(#{c[7]}-H#{@counter})/H#{@counter}" ,c[6]+c[7], "=(#{c[6]+c[7]}-J#{@counter})/J#{@counter}",c[5]], :style => @gridstyle_border
                end                  
              end
            @counter += 1
          end
        end #HONG KONG IF END
          
        #Column Styles
          sheet.col_style 4, @decimal, row_offset: 3
          sheet.col_style 5, @number, row_offset: 3
          sheet.col_style 6, @percent, row_offset: 3
          sheet.col_style 7, @number, row_offset: 3
          sheet.col_style 8, @percent, row_offset: 3
          sheet.col_style 9, @number, row_offset: 3
          sheet.col_style 10, @percent, row_offset: 3
          sheet.col_style 11, @decimal, row_offset: 3
          
      end
    end
    
  end #STYLES END
    
    package.serialize("Rate_Exchange_Analysis.xlsx")
    #send_file("#{Rails.root}/tmp/basic.xlsx", filename: "Basic.xlsx", type: "application/vnd.ms-excel")
  end
end