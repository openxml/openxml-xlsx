module Xlsx
  module Parts
  end
end

Dir.glob("#{File.join(File.dirname(__FILE__), "parts", "*.rb")}").sort.each do |file|
  require file
end
