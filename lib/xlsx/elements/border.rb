module Xlsx
  module Elements
    class Border
      attr_reader :left_component, :right_component, :top_component, :bottom_component, :diagonal_component
      
      def initialize(options={})
        @left_component = options.fetch(:left, BorderComponent.new)
        @right_component = options.fetch(:right, BorderComponent.new)
        @top_component = options.fetch(:top, BorderComponent.new)
        @bottom_component = options.fetch(:bottom, BorderComponent.new)
        @diagonal_component = options.fetch(:diagonal, BorderComponent.new)
      end
      
      def to_xml(xml)
        xml.border do
          left_component.to_xml("left", xml)
          right_component.to_xml("right", xml)
          top_component.to_xml("top", xml)
          bottom_component.to_xml("bottom", xml)
          diagonal_component.to_xml("diagonal", xml)
        end
      end
      
    end
  end
end
