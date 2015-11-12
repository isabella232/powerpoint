require 'erb'

module Powerpoint
  module Slide

    # renders an image for a PPT slide, snapping it to coordinates if passed.
    class ImageContent
      include Powerpoint::Util

      attr_reader :image_name, :coords, :image_path, :id, :idx, :embed_rid
      attr_writer :id, :idx, :embed_rid

      def initialize(options={})
        require_arguments [:image_path], options
        options.each { |k, v| instance_variable_set("@#{k}", v) }
        @idx ||= -1
        @centered ||= false
        if @coords
          @coords = boxed_coordinates(@coords, FastImage.size(@image_path), @centered)
        else
          @coords = {}
        end
        @image_name = File.basename(@image_path)
      end

      def file_type
        File.extname(image_name).gsub('.', '')
      end

      def render
        render_str('picture.xml.erb')
      end
    end
  end
end