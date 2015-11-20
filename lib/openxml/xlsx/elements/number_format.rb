require "openxml/xlsx/elements/implied_number_format"

module OpenXml
  module Xlsx
    module Elements
      class NumberFormat < Struct.new(:format)
        INTEGER = ImpliedNumberFormat.new(1).freeze                  # 0
        DECIMAL = ImpliedNumberFormat.new(2).freeze                  # 0.00
        INTEGER_THOUSANDS = ImpliedNumberFormat.new(3).freeze        # #,##0
        DECIMAL_THOUSANDS = ImpliedNumberFormat.new(4).freeze        # #,##0.00

        INTEGER_PERCENT = ImpliedNumberFormat.new(9).freeze          # 0%
        DECIMAL_PERCENT = ImpliedNumberFormat.new(10).freeze         # 0.00%
        SCIENTIFIC = ImpliedNumberFormat.new(11).freeze              # 0.00E+00
        FRACTION = ImpliedNumberFormat.new(12).freeze                # ?/?
        FRACTION2 = ImpliedNumberFormat.new(13).freeze               # ??/??
        DATE = ImpliedNumberFormat.new(14).freeze                    # mm-dd-yy
        DATE_ALT = ImpliedNumberFormat.new(15).freeze                # d-mmm-yy
        DAY_MONTH = ImpliedNumberFormat.new(16).freeze               # d-mmm
        MONTH_YEAR = ImpliedNumberFormat.new(17).freeze              # mmm-yy
        TIME = ImpliedNumberFormat.new(18).freeze                    # h:mm AM/PM
        TIME_SECONDS = ImpliedNumberFormat.new(19).freeze            # h:mm:ss AM/PM
        TIME_ALT = ImpliedNumberFormat.new(20).freeze                # h:mm
        TIME_SECONDS_ALT = ImpliedNumberFormat.new(21).freeze        # h:mm:ss
        DATETIME = ImpliedNumberFormat.new(22).freeze                # m/d/yy h:mm

        FINANCIAL_INTEGER = ImpliedNumberFormat.new(37).freeze       # #,##0 ;(#,##0)
        FINANCIAL_INTEGER_RED = ImpliedNumberFormat.new(38).freeze   # #,##0 ;[Red](#,##0)
        FINANCIAL_DECIMAL = ImpliedNumberFormat.new(39).freeze       # #,##0.00 ;(#,##0.00)
        FINANCIAL_DECIMAL_RED = ImpliedNumberFormat.new(40).freeze   # #,##0.00 ;[Red](#,##0.00)

        INTERVAL = ImpliedNumberFormat.new(45).freeze                # mm:ss
        INTERVAL_HOURS = ImpliedNumberFormat.new(46).freeze          # [h]:mm:ss
        TIMESTAMP = ImpliedNumberFormat.new(47).freeze               # mmss.0
        # SCIENTIFIC_ALT = ImpliedNumberFormat.new(48).freeze        # ##0.0E+0
        # UNKONWN = ImpliedNumberFormat.new(49).freeze               # @



        def to_xml(id, xml)
          xml.numFmt(numFmtId: id, formatCode: format)
        end

      end
    end
  end
end
