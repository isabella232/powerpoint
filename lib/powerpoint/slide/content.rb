require 'erb'

module Powerpoint
  module Slide

    # a modular slide with a title and any number of other elements, which are
    # each individually rendered and added to the slide xml
    class DefaultSlide
      include Powerpoint::Util

      attr_reader :elements, :title, :slide_layout, :content_xml, :images

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

        @content_xml = gen_xml
      end

      def file_types
        types = elements.map do |element|
          element.file_type if element.respond_to?(:file_type)
        end
        types.compact.uniq
      end

      def gen_xml
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

    # renders an image for a PPT slide, snapping it to coordinates if passed.
    class ImageContent
      include Powerpoint::Util

      attr_reader :image_name, :coords, :image_path, :id, :idx, :embed_rid
      attr_writer :id, :idx, :embed_rid

      def initialize(options={})
        require_arguments [:image_path], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @idx = -1 unless @idx
        @centered = false unless @centered
        if !@coords
          @coords = {}
        else
          @coords = boxed_coordinates(@coords, FastImage.size(@image_path), @centered)
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

    class TableContent
      include Powerpoint::Util

      attr_reader :content, :id, :idx
      attr_writer :id, :idx

      def initialize(options={})
        require_arguments [:content, :idx], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}

      end

      def render
        render_str('table.xml.erb')
      end
    end

    # Allows us to use formatting inside of strings
    #
    # Usage:
    # Text.new('my string', size: 1200, color: 'ab041g') + ' has ' + Text.new('sizes', size: 1600)
    #
    #
    class Text
      include Powerpoint::Util

      attr_reader :values, :next, :options
      attr_writer :next

      def initialize(values, options={})
        # require_arguments [], options
        options.each { |k, v| instance_variable_set("@#{k}", v) }
        if values.instance_of? String
          @values = [values]
        else
          @values = values
        end

        @options = options
        @bullets = false unless @bullets
      end

      def + (other)
        if other.instance_of? String
          if @options.has_key?(:color) || @options.has_key?(:size)
            self.copy(Text.new(other))
          else
            Text.new(@values + [other], options)
          end
        else
          self.copy(other.copy(nil))
        end
      end

      def copy(new_next)
        next_opt = @options.clone
        if next_opt[:next] != nil
          next_opt[:next] = next_opt[:next].copy(new_next)
        elsif new_next != nil
          next_opt[:next] = new_next
        end

        Text.new(values, next_opt)
      end

      def to_str
        next_str = @next == nil ? 'nil' : @next.to_str
        "Text['#{values.join("', '")}'] ~> #{next_str}"
      end

      def render(bullets = false)

        header = '<a:p>
                    <a:pPr marL="0" indent="0">
                      <a:spcBef>
                        <a:spcPts val="0"/>
                      </a:spcBef>' +
                      (bullets ? '' : '<a:buNone/>') +
                    '</a:pPr>'

        footer = '</a:p>'

        tmp = self
        output = header
        while tmp != nil
          output += tmp.render_elem
          tmp = tmp.next
        end
        output + footer
      end

      def render_elem
        sz = (options[:size] == nil) ? ' ' : " sz=\"#{options[:size]}\""
        color = (options[:color] == nil) ? '' : "<a:solidFill>
                                                  <a:srgbClr val=\"#{options[:color]}\"/>
                                                 </a:solidFill>"

        before = '<a:r><a:rPr dirty="0" lang="en-US"' + sz + ' smtClean="0">' + color + '</a:rPr><a:t>'
        after = '</a:t></a:r>'

        "#{before}#{values.join('')}#{after}"
      end
    end

    # renders a text block element
    # takes either Powerpoint::Slide::Text or strings as content
    class TextContent
      include Powerpoint::Util

      attr_reader :content, :id, :idx, :bullets
      attr_writer :id, :idx

      def initialize(options={})
        require_arguments [:content, :idx], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
        @bullets = false unless @bullets
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
