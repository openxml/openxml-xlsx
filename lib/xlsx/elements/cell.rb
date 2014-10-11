module Xlsx
  module Elements
    class Cell
      attr_reader :row, :column, :value, :type, :style, :formula
      
      def initialize(row, options={})
        @row = row
        @column = options.fetch(:column)
        @value = options[:value]
        @type = :general
        if value.is_a? String
          @type = :string
          @string_reference = package.string_ref(value)
        end
        @style = package.style_ref(options[:style]) if options.key? :style
        @formula = options[:formula]
      end
      
      def id
        "#{column_letter}#{row.number}"
      end
      
      def column_letter
        bytes = []
        remaining = column
        while remaining > 0
          bytes.unshift (remaining - 1) % 26 + 65
          remaining = (remaining - 1) / 26
        end
        bytes.pack "c*"
      end
      
      def worksheet
        row.worksheet
      end
      
      def workbook
        worksheet.workbook
      end
      
      def package
        workbook.package
      end
      
      def to_xml(xml)
        attributes = {"r" => id}
        attributes.merge!("t" => TYPE_MAP[type]) unless type == :general
        attributes.merge!("s" => style) if style
        
        value = self.value
        value = string_reference if type == :string
        
        xml.c(attributes) do
          xml.v value if value
          xml.f formula if formula
        end
      end
      
    private
      attr_reader :string_reference
      
      TYPE_MAP = {
        general: nil,
        string: "s" }.freeze
      
    end
  end
end
