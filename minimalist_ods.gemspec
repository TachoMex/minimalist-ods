# frozen_string_literal: true

Gem::Specification.new do |spec|
  spec.name          = "minimalist_ods"
  spec.version       = "0.2.0"
  spec.authors       = ["Gilberto Vargas"]
  spec.email         = ["tachoguitar@gmail.com"]
  spec.summary       = %q{A minimalist ODS generator}
  spec.description   = %q{A minimalist ODS generator written in Ruby}
  spec.homepage      = "https://github.com/TachoMex/minimalist-ods"
  spec.license       = "MIT"

  spec.files         = Dir["lib/*.rb"]
  spec.require_paths = ["lib"]

  spec.add_dependency "rubyzip"
end
