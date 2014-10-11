module Xlsx
  module Parts
    class Rels < BasePart
      attr_reader :relationships

      def initialize(defaults=[])
        @relationships = []
        Array(defaults).each do |default|
          add_relationship(*default.values_at("Type", "Target", "Id"))
        end
      end

      def add_relationship(type, target, id=nil)
        relationships.push Xlsx::Elements::Relationship.new(type, target, id)
      end

      def to_xml
        build_standalone_xml do |xml|
          xml.Relationships(xmlns: "http://schemas.openxmlformats.org/package/2006/relationships") do
            relationships.each do |rel|
              rel.to_xml(xml)
            end
          end
        end
      end

    end
  end
end
