# This file is auto-generated from the current state of the database. Instead
# of editing this file, please use the migrations feature of Active Record to
# incrementally modify your database, and then regenerate this schema definition.
#
# Note that this schema.rb definition is the authoritative source for your
# database schema. If you need to create the application database on another
# system, you should be using db:schema:load, not running all the migrations
# from scratch. The latter is a flawed and unsustainable approach (the more migrations
# you'll amass, the slower it'll run and the greater likelihood for issues).
#
# It's strongly recommended that you check this file into your version control system.

ActiveRecord::Schema.define(version: 20161021195449) do

  # These are extensions that must be enabled in order to support this database
  enable_extension "plpgsql"

  create_table "azerbaijan_pulls", force: :cascade do |t|
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end

  create_table "country_mappings", force: :cascade do |t|
    t.string "currency"
    t.string "country"
    t.string "country_code"
  end

  create_table "excel_reports", force: :cascade do |t|
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end

  create_table "export_vols", force: :cascade do |t|
    t.string  "month"
    t.integer "year"
    t.decimal "export_vol"
    t.string  "country"
  end

  create_table "imf_data", force: :cascade do |t|
    t.datetime "date"
    t.string   "currency_name"
    t.decimal  "rate"
  end

  create_table "outbound_gbs_databases", force: :cascade do |t|
    t.string   "orig"
    t.string   "dest"
    t.string   "ob_ind"
    t.datetime "month"
    t.integer  "year"
    t.string   "mail_class_code"
    t.integer  "pieces"
  end

end
