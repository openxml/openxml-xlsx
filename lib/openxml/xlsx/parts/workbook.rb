module OpenXml
  module Xlsx
    module Parts
      class Workbook < OpenXml::Part
        attr_reader :package, :worksheets, :tables, :defined_names

        def initialize(package)
          @package = package
          @worksheets = []
          @tables = []
          @defined_names =[]
          add_worksheet
        end

        def add_worksheet
          worksheet = Worksheet.new(self, worksheets.length + 1)
          package.xl_rels.add_relationship(
            REL_WORKSHEET,
            "worksheets/sheet#{worksheet.index}.xml",
            "rId#{worksheet.index}")
          package.add_part "xl/worksheets/_rels/sheet#{worksheet.index}.xml.rels", worksheet.rels
          package.add_part "xl/worksheets/sheet#{worksheet.index}.xml", worksheet
          worksheets.push worksheet
        end

        def add_table(table)
          package.content_types.add_override "/xl/tables/#{table.filename}", TYPE_TABLE
          package.add_part "xl/tables/#{table.filename}", table
          tables.push table
        end

        def add_defined_names(*defined_names)
          defined_names.flatten.each do |attributes|
            add_defined_name attributes
          end
        end

        def add_defined_name(attributes)
          defined_names.push Xlsx::Elements::DefinedName.new(attributes[:name], attributes[:formula])
        end

        def to_xml
          build_standalone_xml do |xml|
            xml.workbook(root_namespaces) {
              xml.bookViews {
                xml.workbookView
              }
              xml.sheets { worksheets.each { |worksheet|
                xml.sheet(
                  "name" => worksheet.name,
                  "sheetId" => worksheet.index,
                  "r:id" => "rId#{worksheet.index}")
              } }
              xml.definedNames do
                defined_names.each { |defined_name| defined_name.to_xml(xml) }
              end if defined_names.any?
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
end
