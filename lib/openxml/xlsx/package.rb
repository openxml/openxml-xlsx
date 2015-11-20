require "openxml/package"

module OpenXml
  module Xlsx
    class Package < OpenXml::Package
      attr_reader :xl_rels,
                  :shared_strings,
                  :stylesheet,
                  :workbook


      content_types do
        override "/xl/workbook.xml", TYPE_WORKBOOK
        override "/xl/worksheets/sheet1.xml", TYPE_WORKSHEET
        override "/xl/sharedStrings.xml", TYPE_SHARED_STRINGS
        override "/xl/styles.xml", TYPE_STYLES
      end


      def initialize
        super
        rels.add_relationship REL_DOCUMENT, "xl/workbook.xml"

        @xl_rels = OpenXml::Parts::Rels.new([
          { "Type" => REL_SHARED_STRINGS, "Target" => "sharedStrings.xml" },
          { "Type" => REL_STYLES, "Target" => "styles.xml" }
        ])
        @shared_strings = Xlsx::Parts::SharedStrings.new
        @stylesheet = Xlsx::Parts::Stylesheet.new
        @workbook = Xlsx::Parts::Workbook.new(self)

        # docProps/app.xml
        # docProps/core.xml
        add_part "xl/_rels/workbook.xml.rels", xl_rels
        # xl/calcChain.xml
        add_part "xl/sharedStrings.xml", shared_strings
        add_part "xl/styles.xml", stylesheet
        # xl/theme/theme1.xml
        add_part "xl/workbook.xml", workbook
      end

      def string_ref(string)
        shared_strings.reference_of(string)
      end

      def style_ref(style)
        stylesheet.reference_of(style)
      end

    end
  end
end
