module Xlsx
  module Elements
    class PatternFill < Struct.new(:type, :foreground_color, :background_color)
      
      def to_xml(xml)
        xml.fill do
          xml.patternFill(patternType: type) do
            if type.to_s == "solid"
              foreground_color.to_xml("fgColor", xml)
              background_color.to_xml("bgColor", xml)
            end
          end
        end
      end
      
    end
  end
end
