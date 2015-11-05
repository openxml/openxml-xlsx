module Xlsx
  module Elements
    class Row
      attr_reader :worksheet, :number, :height, :hidden, :cells

      def initialize(worksheet, options={})
        @worksheet = worksheet
        @number = options.fetch(:number)
        @height = options[:height]
        @hidden = options[:hidden]
        @cells = []

        Array(options[:cells]).each do |attributes|
          add_cell attributes
        end
      end

      def add_cell(attributes)
        cells.push Xlsx::Elements::Cell.new(self, attributes)
      end

      def to_xml(xml)
        attributes = {"r" => number}
        attributes.merge!("ht" => height, "customHeight" => 1) if height
        attributes.merge!("hidden" => 1) if hidden
        xml.row(attributes) do
          cells.each do |cell|
            cell.to_xml(xml)
          end
        end
      end

    end
  end
end
