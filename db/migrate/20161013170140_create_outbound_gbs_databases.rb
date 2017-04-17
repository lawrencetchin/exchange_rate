class CreateOutboundGbsDatabases < ActiveRecord::Migration[5.0]
  def change
    create_table :outbound_gbs_databases do |t|
        t.string  :orig
        t.string  :dest
        t.string  :ob_ind
        t.timestamp :month
        t.integer :year
        t.string  :mail_class_code
        t.integer :pieces
        
    end
  end
end
