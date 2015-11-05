module Xlsx
  module Elements
    class DefinedName < Struct.new(:name, :formula)

      def to_xml(xml)
        xml.definedName "#{formula}", name: name
      end

    end
  end
end
