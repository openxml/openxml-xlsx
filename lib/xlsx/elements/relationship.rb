require "securerandom"

module Xlsx
  module Elements
    class Relationship < Struct.new(:type, :target, :id)
      
      def initialize(type, target, id=nil)
        id ||= "R#{SecureRandom.hex}"
        super type, target, id
      end
      
      def to_xml(xml)
        xml.Relationship("Id" => id, "Target" => target, "Type" => type)
      end
      
    end
  end
end
