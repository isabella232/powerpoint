require 'erb'

module Powerpoint
  module Slide


    # renders a text block element
    # takes either Powerpoint::Slide::Text or strings as content
    class TextContent
      include Powerpoint::Util

      attr_reader :content, :id, :idx, :bullets
      attr_writer :id, :idx

      def initialize(options={})
        require_arguments [:content, :idx], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @bullets ||= false
        @content.map! do |content|
          if content.instance_of? String
            Text.new(content)
          else
            content
          end
        end
      end

      def render
        render_str('text.xml.erb')
      end
    end
  end
end
