require "open_xml_package"

module Xlsx
  class Package
    attr_reader :content_types,
                :global_rels,
                :rels,
                :shared_strings,
                :workbook

    def initialize
      @content_types = Xlsx::Parts::ContentTypes.new(nil, {
        "/xl/workbook.xml" => TYPE_WORKBOOK,
        "/xl/worksheets/sheet1.xml" => TYPE_WORKSHEET,
        "/xl/sharedStrings.xml" => TYPE_SHARED_STRINGS
      })
      @global_rels = Xlsx::Parts::Rels.new([
        { "Type" => REL_DOCUMENT, "Target" => "xl/workbook.xml" },
      ])
      @rels = Xlsx::Parts::Rels.new([
        { "Type" => REL_SHARED_STRINGS, "Target" => "sharedStrings.xml" },
      ])
      @shared_strings = Xlsx::Parts::SharedStrings.new
      @workbook = Xlsx::Parts::Workbook.new(self)
    end
    
    def string_reference(string)
      shared_strings.reference_of(string)
    end

    def save(path)
      package = OpenXmlPackage.new
      package.add_part "[Content_Types].xml", content_types.read
      package.add_part "_rels/.rels", global_rels.read
      # docProps/app.xml
      # docProps/core.xml
      package.add_part "xl/_rels/workbook.xml.rels", rels.read
      # xl/calcChain.xml
      package.add_part "xl/sharedStrings.xml", shared_strings.read
      # xl/styles.xml
      # xl/theme/theme1.xml
      package.add_part "xl/workbook.xml", workbook.read
      workbook.worksheets.each do |worksheet|
        package.add_part "xl/worksheets/sheet#{worksheet.index}.xml", worksheet.read
      end
      package.write_to path
    end

  end
end
