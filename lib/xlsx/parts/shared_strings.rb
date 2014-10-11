module Xlsx
  module Parts
    class SharedStrings < BasePart
      attr_reader :strings

      def initialize
        @strings = []
      end

      def reference_of(string)
        Xlsx.index!(strings, string)
      end

      def to_xml
        build_standalone_xml do |xml|
          xml.sst(xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", uniqueCount: strings.length) do
            strings.each do |string|
              xml.si { xml.t(string) }
            end
          end
        end
      end

    end
  end
end
