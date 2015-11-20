module OpenXml
  module Xlsx
    module Elements
      class Alignment < Struct.new(:horizontal, :vertical, :indent, :wrapText)

        def attributes
          {}.tap do |attrs|
            attrs[:horizontal] = horizontal if horizontal
            attrs[:vertical] = vertical if vertical
            attrs[:indent] = indent if indent
            attrs[:wrapText] = wrapText if wrapText
          end
        end

        def to_xml(xml)
          xml.alignment(attributes)
        end

      end
    end
  end
end
