module OpenXml
  module Xlsx
    module Elements
    end
  end
end

Dir.glob("#{File.join(File.dirname(__FILE__), "elements", "*.rb")}").each do |file|
  require file
end
