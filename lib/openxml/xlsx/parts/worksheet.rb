module OpenXml
  module Xlsx
    module Parts
      class Worksheet < OpenXml::Part
        attr_reader :workbook, :index, :rows, :tables, :rels, :cell_ranges

        def initialize(workbook, index)
          @workbook = workbook
          @index = index
          @rows = []
          @tables = []
          @cell_ranges = []
          @rels = OpenXml::Parts::Rels.new
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

        def merge_cells(*ranges)
          ranges.each { |range| cell_ranges.push range }
        end

        def add_table(id, name, ref, columns)
          table = Xlsx::Parts::Table.new(id, name, ref, columns)
          rels.add_relationship(REL_TABLE, "../tables/#{table.filename}")
          workbook.add_table table
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
              xml.mergeCells(count: merge_cells.size) do
                cell_ranges.each { |range| xml.mergeCell ref: range }
              end if cell_ranges.any?
              xml.pageMargins(left: 0.75, right: 0.75, top: 1, bottom: 1, header: 0.5, footer: 0.5)
              xml.pageSetup(orientation: "portrait", horizontalDpi: 4294967292, verticalDpi: 4294967292)
              xml.tableParts(count: tables.count) do
                tables.each do |rel|
                  xml.tablePart("r:id" => rel.id)
                end
              end if tables.any?
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

        def tables
          rels.select { |rel| rel.type == REL_TABLE }
        end

      end
    end
  end
end
