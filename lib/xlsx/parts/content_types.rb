module Xlsx
  module Parts
    class ContentTypes < BasePart
      attr_reader :defaults, :overrides

      def initialize(defaults={}, overrides={})
        @defaults = REQUIRED_DEFAULTS.merge(defaults || {})
        @overrides = overrides || {}
      end

      def set_default(extension, content_type)
        defaults[extension] = content_type
      end

      def set_override(part_name, content_type)
        overrides[part_name] = content_type
      end

      def to_xml
        build_standalone_xml do |xml|
          xml.Types(xmlns: "http://schemas.openxmlformats.org/package/2006/content-types") do
            defaults.each do |extension, content_type|
              xml.Default("Extension" => extension, "ContentType" => content_type)
            end
            overrides.each do |part_name, content_type|
              xml.Override("PartName" => part_name, "ContentType" => content_type)
            end
          end
        end
      end

    private

      REQUIRED_DEFAULTS = {
        "xml" => "application/xml",
        "rels" => "application/vnd.openxmlformats-package.relationships+xml" }.freeze

    end
  end
end
