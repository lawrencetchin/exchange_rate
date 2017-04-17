class AddColumnCountryToExportVol < ActiveRecord::Migration[5.0]
  def change
    add_column :export_vols,  :country, :string
  end
end
