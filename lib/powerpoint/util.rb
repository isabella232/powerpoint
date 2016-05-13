module Powerpoint
  module Util
    extend self

    def pixle_to_pt(px)
      px * 12700
    end

    def render_view(template_name, path)
      view_contents = read_template(template_name)
      renderer = ERB.new(view_contents)
      data = renderer.result(binding)

      File.open(path, 'w') { |f| f << data }
    end

    def read_template(filename)
      File.read("#{Powerpoint::VIEW_PATH}/#{filename}")
    end

    def require_arguments(required_argements, argements)
      raise ArgumentError unless required_argements.all? {|required_key| argements.keys.include? required_key}
    end

    def copy_media(extract_path, image_path)
      image_name = File.basename(image_path)
      dest_path = "#{extract_path}/ppt/media/#{image_name}"
      FileUtils.copy_file(image_path, dest_path) unless File.exist?(dest_path)
    end

    def encode_xml(text)
      text.to_s.encode(xml: :text)
    end

    def render_str(template)
      renderer = ERB.new(read_template(template))
      renderer.result(binding)
    end

    def boxed_coordinates(template_coords, dimensions, centered = false)
      default_width = template_coords[:cx]
      default_height = template_coords[:cy]

      return template_coords unless dimensions

      image_width, image_height = dimensions.map {|d| pixle_to_pt(d)}

      w_ratio = default_width / image_width.to_f
      h_ratio = default_height / image_height.to_f

      # always take lower ratio so that we don't overshoot the
      # template size in either direction
      ratio = [w_ratio, h_ratio].min

      new_width = (image_width.to_f * ratio).round
      new_height = (image_height.to_f * ratio).round

      # adjust start positions if we are centering
      ycoord = template_coords[:y] + (centered ? (template_coords[:cy] / 2) - (new_height / 2) : 0)
      xcoord = template_coords[:x] + (centered ? (template_coords[:cx] / 2) - (new_width / 2) : 0)

      {x: xcoord, y: ycoord, cx: new_width, cy: new_height}
    end

    def vertically_grouped_coordinates(template_coords, image_dimensions)
      dimensions = image_dimensions.map { |dims| dims.map { |d| pixle_to_pt(d) } }

      total_img_height = dimensions.map { |d| d[1] }.reduce(:+)
      total_img_width = dimensions.map { |d| d[0] }.max

      h_ratio = template_coords[:cy] / total_img_height.to_f
      w_ratio = template_coords[:cx] / total_img_width.to_f

      # always take lower ratio so that we don't overshoot the
      # template size in either direction
      ratio = [w_ratio, h_ratio].min

      h_offset = template_coords[:y]

      dimensions.map do |w, h|
        next_width = (w.to_f * ratio).round
        next_height = (h.to_f * ratio).round

        out = {x: template_coords[:x], y: h_offset, cx: next_width, cy: next_height}
        h_offset += next_height

        out
      end
    end

  end
end
