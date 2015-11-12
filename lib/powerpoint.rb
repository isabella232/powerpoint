require 'powerpoint/version'
require 'powerpoint/util'
require 'powerpoint/slide/intro'
require 'powerpoint/slide/content'
require 'powerpoint/compression'
require 'powerpoint/presentation'
require 'pry'

module Powerpoint
  ROOT_PATH = File.expand_path("../..", __FILE__)
  TEMPLATE_PATH = "#{ROOT_PATH}/template/"
  VIEW_PATH = "#{ROOT_PATH}/lib/powerpoint/views/"
end
