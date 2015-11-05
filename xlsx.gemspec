# coding: utf-8
lib = File.expand_path("../lib", __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require "xlsx/version"

Gem::Specification.new do |spec|
  spec.name          = "openxml-xlsx"
  spec.version       = Xlsx::VERSION
  spec.authors       = ["Bob Lail"]
  spec.email         = ["bob.lail@cph.org"]

  spec.description   = %q{Create Microsoft Excel (.xlsx) files.}
  spec.summary       = %q{Using a simple API, create xlsx files programmatically}
  spec.license       = "MIT"
  spec.homepage      = "https://github.com/concordia-publishing-house/xlsx"

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.required_ruby_version = "~> 2.0"
  spec.add_dependency "nokogiri"
  spec.add_dependency "open_xml_package", "0.2.0.beta1"

  spec.add_development_dependency "pry"
  spec.add_development_dependency "rspec"
  spec.add_development_dependency "rake"
  spec.add_development_dependency "rr"
  spec.add_development_dependency "simplecov"
  spec.add_development_dependency "timecop"
end
