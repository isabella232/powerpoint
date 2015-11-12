require 'erb'

module Powerpoint
  module Slide
    # a modular slide with a title and any number of other elements, which are
    # each individually rendered and added to the slide xml
    class DefaultSlide
      include Powerpoint::Util

      attr_reader :elements, :title, :slide_layout, :images

      def initialize(options={})
        require_arguments [:elements, :title, :slide_layout], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @images = elements.select { |image| image.respond_to?(:embed_rid) }

        @images.each_with_index do |image, i|
          image.embed_rid = "rId#{i+2}"
        end

        @elements.each_with_index do |elem, i|
          elem.id = 10 + i
        end
      end

      def file_types
        types = elements.map do |element|
          element.file_type if element.respond_to?(:file_type)
        end
        types.compact.uniq
      end

      def content_xml
        xmls = elements.map { |e| e.render }
        xmls.join('')
      end

      def render_rels
        render_str('default_slide_rel.xml.erb')
      end

      def render_slide
        render_str('default_slide.xml.erb')
      end

      def save(extract_path, index)
        images.each do |image|
          copy_media(extract_path, image.image_path)
        end

        File.open("#{extract_path}/ppt/slides/_rels/slide#{index}.xml.rels", 'w') do |f|
          f << render_rels
        end

        File.open("#{extract_path}/ppt/slides/slide#{index}.xml", 'w') do |f|
          f << render_slide
        end

      end
    end
  end
end