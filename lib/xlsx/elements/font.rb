module Xlsx
  module Elements
    class Font < Struct.new(:name, :size)
      
      def to_xml(xml)
        xml.font do
          xml.sz val: size
          xml.name val: name
        end
      end
      
    end
  end
end
