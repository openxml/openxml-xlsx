module OpenXml
  module Xlsx
    module Elements
      class BorderComponent < Struct.new(:style, :color)

        def to_xml(name, xml)
          if style && color
            xml.public_send(name, style: style) do
              color.to_xml("color", xml)
            end
          elsif style
            xml.public_send(name, style: style)
          else
            xml.public_send(name)
          end
        end

      end
    end
  end
end
