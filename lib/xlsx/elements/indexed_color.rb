module Xlsx
  module Elements
    class IndexedColor < Struct.new(:indexed)

      def to_xml(name, xml)
        xml.public_send(name, indexed: indexed)
      end

    end
  end
end
