class CreateExcelReports < ActiveRecord::Migration[5.0]
  def change
    create_table :excel_reports do |t|

      t.timestamps
    end
  end
end
