module Xlsx
  module Elements
    class Border
      
      def to_xml(xml)
        xml.border do
          xml.left
          xml.right
          xml.top
          xml.bottom
          xml.diagonal
        end
      end
      
    end
  end
end
