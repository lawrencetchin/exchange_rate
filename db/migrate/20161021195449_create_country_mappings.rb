class CreateCountryMappings < ActiveRecord::Migration[5.0]
  def change
    create_table :country_mappings do |t|
      t.string  :currency
      t.string  :country
      t.string  :country_code
    end
  end
end
