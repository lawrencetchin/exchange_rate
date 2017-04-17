class CreateExportVols < ActiveRecord::Migration[5.0]
  def change
    create_table :export_vols do |t|
      t.string    :month
      t.integer   :year
      t.decimal   :export_vol
    end
  end
end
