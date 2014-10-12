module Xlsx
  module Parts
    class Table < BasePart
      attr_reader :id, :name, :ref, :columns

      def initialize(id, name, ref, columns)
        @id = id
        @name = name
        @ref = ref
        @columns = columns
      end

      def filename
        "table#{id}.xml"
      end

      def to_xml
        build_standalone_xml do |xml|
          xml.table(xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            id: id, name: name, displayName: name, ref: ref, totalsRowShown: 0) do
            xml.autoFilter ref: ref
            xml.tableColumns(count: columns.length) do
              columns.each_with_index do |column, index|
                xml.tableColumn(id: index + 1, name: column.name)
              end
            end
            xml.tableStyleInfo(name: "TableStyleLight6", showFirstColumn: 0, showLastColumn: 0, showRowStripes: 1, showColumnStripes: 0)
          end
        end
      end

    end
  end
end
