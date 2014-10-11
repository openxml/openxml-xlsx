require "xlsx/elements"
require "xlsx/package"
require "xlsx/parts"

module Xlsx
  REL_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument".freeze
  REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet".freeze
  REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles".freeze
  REL_SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings".freeze
  REL_CALC_CHAIN = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain".freeze
  REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme".freeze
  
  TYPE_WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".freeze
  TYPE_WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml".freeze
  TYPE_THEME = "application/vnd.openxmlformats-officedocument.theme+xml".freeze
  TYPE_STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml".freeze
  TYPE_SHARED_STRINGS =  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml".freeze
  TYPE_CALC_CHAIN = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml".freeze
  TYPE_CORE_PROPS = "application/vnd.openxmlformats-package.core-properties+xml".freeze
  TYPE_APP_PROPS = "application/vnd.openxmlformats-officedocument.extended-properties+xml".freeze
  
  def self.index!(collection, item)
    collection.index(item) || collection.push(item).length - 1
  end
  
end
