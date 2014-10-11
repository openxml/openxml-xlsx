require "spec_helper"

describe Xlsx::Package do
  attr_reader :package

  context "when starting a new package" do
    before(:each) do
      @package = described_class.new
    end

    it "should create the content types part" do
      expect(package.content_types).to be_instance_of(Xlsx::Parts::ContentTypes)
    end

    it "should create the workbook part" do
      expect(package.workbook).to be_instance_of(Xlsx::Parts::Workbook)
    end
    
    it "should create the global rels part" do
      expect(package.global_rels).to be_instance_of(Xlsx::Parts::Rels)
    end
    
    it "should create the _rels part" do
      expect(package.rels).to be_instance_of(Xlsx::Parts::Rels)
    end
    
    it "should create the shared_strings part" do
      expect(package.shared_strings).to be_instance_of(Xlsx::Parts::SharedStrings)
    end
  end

  context "when saving a package" do
    before(:each) do
      @package = described_class.new
    end

    it "should write to a OpenXmlPackage" do
      path = "some_file.docx"
      mock.instance_of(OpenXmlPackage).write_to(path)
      package.save(path)
    end
  end

end
