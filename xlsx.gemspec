# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'xlsx/version'

Gem::Specification.new do |gem|
  gem.name          = "xlsx"
  gem.version       = Xlsx::VERSION
  gem.authors       = ["Bob Lail"]
  gem.email         = ["bob.lail@cph.org"]
  gem.description   = %q{Create Microsoft Excel (.xlsx) files.}
  gem.summary       = %q{Using a simple API, create xlsx files programmatically}
  gem.license       = "MIT"
  gem.homepage      = "https://github.com/concordia-publishing-house/xlsx"
  gem.required_ruby_version = "~> 2.0"

  gem.add_dependency "nokogiri"
  gem.add_dependency "open_xml_package", "0.1.0.beta1"

  gem.add_development_dependency "pry"
  gem.add_development_dependency "rspec"
  gem.add_development_dependency "rake"
  gem.add_development_dependency "rr"
  gem.add_development_dependency "simplecov"
  gem.add_development_dependency "timecop"

  gem.files         = `git ls-files`.split($/)
  gem.test_files    = Dir.glob('test/**/*_test.rb')
  gem.require_paths = ["lib"]
end
