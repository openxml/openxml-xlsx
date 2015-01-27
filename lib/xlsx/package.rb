require "open_xml_package"

module Xlsx
  class Package
    attr_reader :content_types,
                :global_rels,
                :xl_rels,
                :shared_strings,
                :stylesheet,
                :workbook

    def initialize
      @content_types = Xlsx::Parts::ContentTypes.new(nil, {
        "/xl/workbook.xml" => TYPE_WORKBOOK,
        "/xl/worksheets/sheet1.xml" => TYPE_WORKSHEET,
        "/xl/sharedStrings.xml" => TYPE_SHARED_STRINGS,
        "/xl/styles.xml" => TYPE_STYLES
      })
      @global_rels = Xlsx::Parts::Rels.new([
        { "Type" => REL_DOCUMENT, "Target" => "xl/workbook.xml" },
      ])
      @xl_rels = Xlsx::Parts::Rels.new([
        { "Type" => REL_SHARED_STRINGS, "Target" => "sharedStrings.xml" },
        { "Type" => REL_STYLES, "Target" => "styles.xml" }
      ])
      @shared_strings = Xlsx::Parts::SharedStrings.new
      @stylesheet = Xlsx::Parts::Stylesheet.new
      @workbook = Xlsx::Parts::Workbook.new(self)
    end

    def string_ref(string)
      shared_strings.reference_of(string)
    end

    def style_ref(style)
      stylesheet.reference_of(style)
    end

    def to_stream
      package.to_stream
    end

    def save(path)
      package.write_to path
    end

  private

    def package
      OpenXmlPackage.new.tap do |package|
        package.add_part "[Content_Types].xml", content_types.read
        package.add_part "_rels/.rels", global_rels.read
        # docProps/app.xml
        # docProps/core.xml
        package.add_part "xl/_rels/workbook.xml.rels", xl_rels.read
        # xl/calcChain.xml
        package.add_part "xl/sharedStrings.xml", shared_strings.read
        package.add_part "xl/styles.xml", stylesheet.read
        workbook.tables.each do |table|
          package.add_part "xl/tables/#{table.filename}", table.read
        end
        # xl/theme/theme1.xml
        package.add_part "xl/workbook.xml", workbook.read
        workbook.worksheets.each do |worksheet|
          package.add_part "xl/worksheets/_rels/sheet#{worksheet.index}.xml.rels", worksheet.rels.read if worksheet.rels.any?
          package.add_part "xl/worksheets/sheet#{worksheet.index}.xml", worksheet.read
        end
      end
    end

  end
end
