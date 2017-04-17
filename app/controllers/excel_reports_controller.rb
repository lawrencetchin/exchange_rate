class ExcelReportsController < ApplicationController
  def index
    
    package = Axlsx::Package.new
    workbook = package.workbook
    @rates = ImfDatum.where('date >= ?','2016-10-01').pluck('date, currency_name, rate')
    
    workbook.styles do |s|
      @heading = s.add_style alignment: {horizontal: :center}, b: true, sz: 18, bg_color: "0066CC", fg_color: "FF"
      @center = s.add_style alignment: {horizontal: :center}, fg_color: "0000FF"
    @header = s.add_style alignment: {horizontal: :center}, b: true, sz: 10, bg_color: "C0C0C0"
    
    end
    
    workbook.add_worksheet(name: "Basic work sheet") do |sheet|
      #sheet.add_row
      sheet.add_row ["","Date", "Currency Name", "Rate"], style: @heading
      counter = 0
      @rates.each do |st|
        sheet.add_row [counter+1, st[0], st[1], st[2]]
        counter += 1
      end
      
      
      sheet.row_style 0, @header, col_offset: 1
      sheet.col_style 4, @center, row_offset: 1
    end
    
    
    
    package.serialize("basic.xlsx")
    #send_file("#{Rails.root}/tmp/basic.xlsx", filename: "Basic.xlsx", type: "application/vnd.ms-excel")
  end
  
end
