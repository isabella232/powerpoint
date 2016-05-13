require 'erb'

module Powerpoint
  module Slide

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

        @values = Array(values)
        @options = options
        @bullets = false unless @bullets
      end

      def + (other)
        if other.instance_of? String
          self.copy(Text.new(other))
        else
          self.copy(other.copy(nil))
        end
      end

      def copy(new_next)
        next_opt = @options.clone
        if next_opt[:next]
          next_opt[:next] = next_opt[:next].copy(new_next)
        elsif new_next
          next_opt[:next] = new_next
        end

        Text.new(values, next_opt)
      end

      def to_str
        next_str = @next ? 'nil' : @next.to_str
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
        sz = (options[:size].nil?) ? ' ' : %(sz="#{options[:size]}")
        color = (options[:color].nil?) ? '' : %(<a:solidFill>
                                                  <a:srgbClr val="#{options[:color]}"/>
                                                 </a:solidFill>)
        bold = (options[:bold]) ? 'b="1"' : ''

        before = %(<a:r><a:rPr dirty="0" lang="en-US" #{sz} #{bold} smtClean="0">#{color}</a:rPr><a:t>)
        after = '</a:t></a:r>'

        "#{before}#{values.map { |b| b.encode(:xml => :text) }.join}#{after}"
      end
    end
  end
end