require "date"

module OpenXml
  module Xlsx
    module Elements
      class Cell
        attr_reader :row, :column, :value, :type, :style, :formula

        def initialize(row, options={})
          @row = row
          @column = options.fetch(:column)
          @value = options[:value]
          case value
          when Numeric
            @type = :general
          when Date
            @type = :date
            @serial_date = to_serial_date(value)
          when Time
            @type = :time
            @serial_time = to_serial_time(value)
          else
            @value = value.to_s
            @type = :string
            @string_id = package.string_ref(value)
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
          attributes.merge!("s" => style) if style
          attributes.merge!("t" => "s") if type == :string

          value = self.value
          value = string_id if type == :string
          value = serial_date if type == :date
          value = serial_time if type == :time

          xml.c(attributes) do
            xml.f formula if formula
            xml.v value if value
          end
        end

      private
        attr_reader :string_id, :serial_date, :serial_time

        EXCEL_ANCHOR_DATE = Date.new(1900, 1, 1).freeze
        SECONDS_PER_DAY = 86400.freeze

        def to_serial_date(date)
          # Excel stores dates as the number of days since 1900-Jan-0
          # Excel behaves as if 1900 was a leap year, so the number is
          # generally 1 greater than you would expect.
          # http://www.cpearson.com/excel/datetime.htm
          (date - EXCEL_ANCHOR_DATE).to_i + 2
        end

        def to_serial_time(time)
          date = to_serial_date(time.to_date)

          seconds_since_midnight = time.hour * 3600 + time.min * 60 + time.sec
          date + (seconds_since_midnight.to_f / SECONDS_PER_DAY)
        end

      end
    end
  end
end
