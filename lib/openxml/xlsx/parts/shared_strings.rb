module OpenXml
  module Xlsx
    module Parts
      class SharedStrings < OpenXml::Part
        attr_reader :strings

        def initialize
          @strings = Hash.new { |hash, key| hash[key] = hash.length }
        end

        def reference_of(string)
          strings[string]
        end

        def to_xml
          build_standalone_xml do |xml|
            xml.sst(xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", uniqueCount: strings.length) do
              strings.each do |string, i|
                xml.si { xml.t(string) }
              end
            end
          end
        end

      end
    end
  end
end
