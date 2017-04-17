class CreateImfData < ActiveRecord::Migration[5.0]
  def change
    create_table :imf_data do |t|
      t.timestamp :date
      t.string   :currency_name
      t.decimal   :rate
    end
  end
end
