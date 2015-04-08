module Xlsx
  module Elements
    class Font < Struct.new(:name, :size, :styles)
      
      def to_xml(xml)
        xml.font do
          xml.sz val: size
          xml.name val: name
          if styles
            xml.b if styles[:bold]
            xml.i if styles[:italic]
            xml.u if styles[:underline]
          end
        end
      end
      
    end
  end
end
