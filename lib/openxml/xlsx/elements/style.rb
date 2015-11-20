module OpenXml
  module Xlsx
    module Elements
      class Style < Struct.new(:format_id, :font_id, :fill_id, :border_id, :alignment)

        def initialize(format_id=0, font_id=0, fill_id=0, border_id=0, alignment=nil)
          super format_id || 0, font_id || 0, fill_id || 0, border_id || 0, alignment
        end

        def to_xml(xml)
          attributes = {
            numFmtId: format_id,
            fontId: font_id,
            fillId: fill_id,
            borderId: border_id }
          attributes.merge!(applyNumberFormat: 1) if format_id > 0
          attributes.merge!(applyFont: 1) if font_id > 0
          attributes.merge!(applyFill: 1) if fill_id > 0
          attributes.merge!(applyBorder: 1) if border_id > 0
          attributes.merge!(applyAlignment: 1) if alignment
          xml.xf(attributes) do
            alignment.to_xml(xml) if alignment
          end
        end

      end
    end
  end
end
