module Xlsx
  module Parts
    class Workbook < BasePart
      attr_reader :package, :worksheets

      def initialize(package)
        @package = package
        @worksheets = []
        add_worksheet
      end

      def add_worksheet
        worksheet = Worksheet.new(self, worksheets.length + 1)
        package.rels.add_relationship(
          REL_WORKSHEET,
          "worksheets/sheet#{worksheet.index}.xml",
          "rId#{worksheet.index}")
        worksheets.push worksheet
      end

      def to_xml
        build_standalone_xml do |xml|
          xml.workbook(root_namespaces) {
            xml.sheets { worksheets.each { |worksheet|
              xml.sheet(
                "name" => worksheet.name,
                "sheetId" => worksheet.index,
                "r:id" => "rId#{worksheet.index}")
            } }
          }
        end
      end

    private

      def root_namespaces
        { "xmlns" => "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "xmlns:r" => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' }
      end

    end
  end
end
