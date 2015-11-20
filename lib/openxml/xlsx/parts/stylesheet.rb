module OpenXml
  module Xlsx
    module Parts
      class Stylesheet < OpenXml::Part
        include Xlsx::Elements

        attr_reader :formats, :fonts, :fills, :borders, :styles

        def initialize
          @formats = []
          @fonts = [Font.new("Calibri", 12)]
          @fills = [PatternFill.new("none"), PatternFill.new("gray125")]
          @borders = [Border.new]
          @styles = [Style.new]
        end

        def reference_of(options={})
          case format = options[:format]
          when NumberFormat
            options[:format_id] = Xlsx.index!(formats, format) + NUMBER_FORMAT_START_ID
          when ImpliedNumberFormat
            options[:format_id] = format.id
          end
          options[:font_id] = Xlsx.index!(fonts, options[:font]) if options.key? :font
          options[:fill_id] = Xlsx.index!(fills, options[:fill]) if options.key? :fill
          options[:border_id] = Xlsx.index!(borders, options[:border]) if options.key? :border

          style = Style.new(
            options.fetch(:format_id, 0),
            options.fetch(:font_id, 0),
            options.fetch(:fill_id, 0),
            options.fetch(:border_id, 0),
            options.fetch(:alignment, nil))

          Xlsx.index!(styles, style)
        end

        def to_xml
          build_standalone_xml do |xml|
            xml.styleSheet(xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "xmlns:mc" => "http://schemas.openxmlformats.org/markup-compatibility/2006") do
              xml.numFmts(count: formats.length) do
                formats.each_with_index do |format, index|
                  format.to_xml(index + NUMBER_FORMAT_START_ID, xml)
                end
              end

              xml.fonts(count: fonts.length) do
                fonts.each do |font|
                  font.to_xml(xml)
                end
              end

              xml.fills(count: fills.length) do
                fills.each do |fill|
                  fill.to_xml(xml)
                end
              end

              xml.borders(count: borders.length) do
                borders.each do |border|
                  border.to_xml(xml)
                end
              end

              xml.cellXfs(count: styles.length) do
                styles.each do |style|
                  style.to_xml(xml)
                end
              end
            end
          end
        end

      private
        NUMBER_FORMAT_START_ID = 165.freeze

      end
    end
  end
end
