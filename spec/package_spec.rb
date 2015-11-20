require "spec_helper"

describe OpenXml::Xlsx::Package do
  attr_reader :package

  context "when starting a new package" do
    before(:each) do
      @package = described_class.new
    end

    it "should create the content types part" do
      expect(package.content_types).to be_instance_of(OpenXml::Parts::ContentTypes)
    end

    it "should create the workbook part" do
      expect(package.workbook).to be_instance_of(OpenXml::Xlsx::Parts::Workbook)
    end

    it "should create the global rels part" do
      expect(package.rels).to be_instance_of(OpenXml::Parts::Rels)
    end

    it "should create the _rels part" do
      expect(package.xl_rels).to be_instance_of(OpenXml::Parts::Rels)
    end

    it "should create the shared_strings part" do
      expect(package.shared_strings).to be_instance_of(OpenXml::Xlsx::Parts::SharedStrings)
    end

    it "should create the stylesheet part" do
      expect(package.stylesheet).to be_instance_of(OpenXml::Xlsx::Parts::Stylesheet)
    end
  end

end
