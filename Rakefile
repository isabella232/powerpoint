require "bundler/gem_tasks"
require 'rspec/core/rake_task'
require "csd-gem-helper"

RSpec::Core::RakeTask.new('spec')

# If you want to make this the default task
task :default => :spec

Bundler::GemHelper.install_tasks
ClearstoryData::GemHelper.install_tasks
ClearstoryData::GemHelper.disable_default_release_task
