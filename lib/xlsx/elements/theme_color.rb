module Xlsx
  module Elements
    class ThemeColor < Struct.new(:theme, :tint)
      
      def to_xml(name, xml)
        xml.public_send(name, theme: theme, tint: tint)
      end
      
    end
  end
end
