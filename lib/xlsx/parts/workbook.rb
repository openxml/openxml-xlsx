module Xlsx
  module Parts
    class Workbook < BasePart
      attr_reader :package, :worksheets, :tables

      def initialize(package)
        @package = package
        @worksheets = []
        @tables = []
        add_worksheet
      end

      def add_worksheet
        worksheet = Worksheet.new(self, worksheets.length + 1)
        package.xl_rels.add_relationship(
          REL_WORKSHEET,
          "worksheets/sheet#{worksheet.index}.xml",
          "rId#{worksheet.index}")
        worksheets.push worksheet
      end

      def add_table(table)
        package.content_types.set_override "/xl/tables/#{table.filename}", TYPE_TABLE
        tables.push table
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
