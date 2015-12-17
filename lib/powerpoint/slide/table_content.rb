require 'erb'

module Powerpoint
  module Slide
    class TableContent
      include Powerpoint::Util

      attr_reader :content, :id, :idx, :unstyled
      attr_writer :id, :idx

      def initialize(options={})
        require_arguments [:content, :idx], options
        options.each { |k, v| instance_variable_set("@#{k}", v) }
      end

      def render
        render_str('_table.xml.erb')
      end
    end
  end
end