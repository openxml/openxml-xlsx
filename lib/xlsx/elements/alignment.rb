module Xlsx
  module Elements
    class Alignment < Struct.new(:horizontal, :vertical)

      def to_xml(xml)
        xml.alignment(horizontal: horizontal, vertical: vertical)
      end

    end
  end
end
