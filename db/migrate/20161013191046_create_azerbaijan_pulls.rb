class CreateAzerbaijanPulls < ActiveRecord::Migration[5.0]
  def change
    create_table :azerbaijan_pulls do |t|

      t.timestamps
    end
  end
end
