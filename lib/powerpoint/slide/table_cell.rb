require 'erb'

module Powerpoint
  module Slide
    class TableCell
      include Powerpoint::Util

      attr_reader :horizontal_size, :vertical_size, :data, :right, :bottom

      def initialize(data, options = {})
        @data = data
        @options = options
        @right = []
        @bottom = []
      end

      # === For building stacks of cells easily ===

      def self.row(data_array)
        cells = data_array.map { |d| TableCell.new(d) }
        cells.each_cons(2) { |a, b| a.set_right_neighbors(b) }
        cells[0]
      end

      def self.column(data_array)
        cells = data_array.map { |d| TableCell.new(d) }
        cells.each_cons(2) { |a, b| a.set_bottom_neighbors(b) }
        cells[0]
      end

      def self.weld_horizontal(left, right)
        left_side = left.right_most
        right_side = right.left_most

        if left_side.size == right_side.size

          left_side.zip(right_side).each do |l, r|
            l.set_right_neighbors(r)
          end
        elsif left_side.size == 1
          left_side[0].set_right_neighbors(right_side)
        else
          raise ArgumentError, "left and right side of weld did not match [left: #{left_side.size} < or != right: #{right_side.size}]"
        end

        left
      end

      def self.weld_vertical(top, bottom)
        top_side = top.bottom_most
        bottom_side = bottom.top_most

        if top_side.size == bottom_side.size
          top_side.zip(bottom_side).each do |l, r|
            l.set_bottom_neighbors(r)
          end
        elsif top_side.size == 1
          top_side[0].set_bottom_neighbors(bottom_side)
        else
          raise ArgumentError, "top and bottom side of weld did not match [top: #{top_side.size} < or != bottom: #{bottom_side.size}]"
        end

        top
      end

      # === Renderable cell conversion ===

      def to_renderable(go_down = true, template = read_template('_table_cell.xml.erb'))

        horizontal, vertical = calculate_sizing

        cells =
            (0..vertical-1).map do |i|
              (0..horizontal-1).map do |j|

                options = {}
                if i == 0 && j == 0
                  options[:gridspan] = horizontal >= 2 ? horizontal : false
                  options[:rowspan] = vertical >= 2 ? vertical : false
                  options[:data] = @data

                elsif i == 0
                  options[:h_merge] = 1
                  options[:rowspan] = vertical >= 2 ? vertical : false
                  options[:template] = template
                elsif j == 0
                  options[:gridspan] = horizontal >= 2 ? horizontal : false
                  options[:v_merge] = i
                  options[:template] = template
                else
                  options[:h_merge] = 1
                  options[:v_merge] = i
                  options[:template] = template
                end

                options[:template] = template

                RenderableCell.new(options.merge(@options))
              end
            end

        neighbor_rows = @right.flat_map { |n| n.to_renderable(false, template) }

        # join right neighbor rows with mine if we have neighbors
        cells = cells.zip(neighbor_rows).map { |myrow, nrow| myrow + nrow } unless neighbor_rows.empty?

        return cells if @bottom.empty? || !go_down

        bot_rows = @bottom[0].to_renderable(true, template)

        cells + bot_rows
      end

      # gets the horizontal and vertical size of the single tile in question
      def calculate_sizing
        return all_below.size, all_right.size
      end

      # === Low level mutation methods ===
      # use with care

      def set_right_neighbors(neighbor)
        @right = Array(neighbor)
      end

      def set_bottom_neighbors(neighbor)
        @bottom = Array(neighbor)
      end


      # === Helper methods ===

      def top_most
        @right.empty? ? [self] : [self] + @right[0].top_most
      end

      def left_most
        @bottom.empty? ? [self] : [self] + @bottom[0].left_most
      end

      def top_right
        top_most[-1]
      end

      # all cells directly below the cell in question
      def all_below
        @bottom.empty? ? [self] : @bottom.flat_map { |r| r.all_below }.uniq
      end

      # all cells to the right of the cell in question
      def all_right
        @right.empty? ? [self] : @right.flat_map { |r| r.all_right }.uniq
      end

      def collect_downwards_stay_right
        @bottom.empty? ? [self] : [self] + @bottom[-1].collect_downwards_stay_right
      end

      def collect_rightwards_stay_down
        @right.empty? ? [self] : [self] + @right[-1].collect_rightwards_stay_down
      end

      def bottom_left
        left_most[-1]
      end

      def right_most
        top_right.collect_downwards_stay_right
      end

      def bottom_most
        bottom_left.collect_rightwards_stay_down
      end
    end

    # Slide content object that render stacks of cells into a formatted table
    class EnhancedTableContent
      include Powerpoint::Util

      attr_reader :content, :id, :idx, :color, :unstyled
      attr_writer :id, :idx

      def initialize(options={})
        require_arguments [:table_cells, :idx], options
        options.each { |k, v| instance_variable_set("@#{k}", v) }
      end

      def render
        @content = @table_cells.to_renderable
        render_str('_enhanced_table.xml.erb')
      end
    end

    # Renderable cells are 1:1 with the 'tc' objects that are added to powerpoint xml,
    # but many of them can be generated from a single TableCell object in the case of merged cells
    class RenderableCell
      include Powerpoint::Util

      attr_reader :h_merge, :v_merge, :rowspan, :gridspan, :data, :color

      def initialize(options)
        require_arguments [:template], options
        options.each {|k, v| instance_variable_set("@#{k}", v)}
      end

      def render
        renderer = ERB.new(@template)
        renderer.result(binding)
      end

      def to_s
        "RCell[#{data} (hm:#{@h_merge}, vm:#{@v_merge}, rs:#{@rowspan}, gs: #{gridspan})]"
      end

    end
  end
end
