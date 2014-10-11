module Xlsx
  module Parts
    class Worksheet < BasePart
      attr_reader :workbook, :index, :rows

      def initialize(workbook, index)
        @workbook = workbook
        @index = index
        @rows = []
        @column_widths = {}
      end

      def column_widths(*args)
        return @column_widths if args.none?
        @column_widths = args.first
      end

      def add_rows(*rows)
        rows.flatten.each do |attributes|
          add_row attributes
        end
      end

      def add_row(attributes)
        rows.push Xlsx::Elements::Row.new(self, attributes)
      end

      def to_xml
        build_standalone_xml do |xml|
          xml.worksheet(root_namespaces) do
            xml.sheetViews do
              xml.sheetView(showGridLines: 0, tabSelected: 1, workbookViewId: 0)
            end
            xml.sheetFormatPr(baseColWidth: 10, defaultColWidth: 13.33203125, defaultRowHeight: 20, customHeight: 1)
            xml.cols do
              column_widths.each do |column, width|
                xml.col(min: column, max: column, width: width, customWidth: 1)
              end
            end if column_widths.any?
            xml.sheetData do
              rows.each { |row| row.to_xml(xml) }
            end
            xml.pageMargins(left: 0.75, right: 0.75, top: 1, bottom: 1, header: 0.5, footer: 0.5)
            xml.pageSetup(orientation: "portrait", horizontalDpi: 4294967292, verticalDpi: 4294967292)
          end
        end
      end

      def name
        "Sheet#{index}"
      end

    private

      def root_namespaces
        { "xmlns" => "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
          "xmlns:r" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
          "xmlns:mc" => "http://schemas.openxmlformats.org/markup-compatibility/2006" }
      end

    end
  end
end
