require 'zip/filesystem'
require 'fileutils'
require 'tmpdir'

module Powerpoint
  class Presentation
    include Powerpoint::Util

    attr_reader :slides

    def initialize
      @slides = []
    end

    def add_intro(title, subtitile = nil)
      existing_intro_slide = @slides.select {|s| s.class == Powerpoint::Slide::Intro}[0]
      slide = Powerpoint::Slide::Intro.new(presentation: self, title: title, subtitile: subtitile)
      if existing_intro_slide
        @slides[@slides.index(existing_intro_slide)] = slide 
      else
        @slides.insert 0, slide
      end
    end

    def add_textual_slide(title, content = [])
      slide = Powerpoint::Slide::DefaultSlide.new(
        title: title,
        slide_layout: 2,
        elements: [
          Powerpoint::Slide::TextContent.new(idx: 1, bullets: true, content: content)
      ])

      @slides << slide
    end

    def add_pictorial_slide(title, image_path, coords = {})

      # magic coordinates from the slide layout
      img_coords = { x: 826325, y: 1825625, cx: 10515600, cy: 4351338 }

      slide = Powerpoint::Slide::DefaultSlide.new(
        title: title,
        slide_layout: 2,
        elements: [
          Powerpoint::Slide::ImageContent.new(
            idx: 1,
            image_path: image_path,
            centered: true,
            coords: img_coords)
        ])

      @slides << slide
    end

    def add_text_picture_slide(title, image_path, content = [])

      # magic coordinates from the slide layout
      img_coords = { x: 6172200, y: 1825625, cx: 5181600, cy: 4351338 }

      slide = Powerpoint::Slide::DefaultSlide.new(
        title: title,
        slide_layout: 4,
        elements: [
          Powerpoint::Slide::TextContent.new(idx: 1, content: content),
          Powerpoint::Slide::ImageContent.new(idx: 2, image_path: image_path, coords: img_coords)
        ])

      @slides << slide
    end

    def add_picture_description_slide(title, image_path, content = [])

      # magic coordinates from the slide layout
      img_coords = { x: 5183188, y: 987425, cx: 6172200, cy: 4873625 }

      slide = Powerpoint::Slide::DefaultSlide.new(
        title: title,
        slide_layout: 9,
        elements: [
          Powerpoint::Slide::ImageContent.new(idx: 1, image_path: image_path, coords: img_coords),
          Powerpoint::Slide::TextContent.new(idx: 2, bullets: true, content: content)
        ])

      @slides << slide
    end

    def add_slide(slide)
      @slides << slide
    end


    def save(path)
      Dir.mktmpdir do |dir|
        extract_path = "#{dir}/extract_#{Time.now.strftime("%Y-%m-%d-%H%M%S")}"

        # Copy template to temp path
        FileUtils.copy_entry(TEMPLATE_PATH, extract_path)

        # Remove keep files
        Dir.glob("#{extract_path}/**/.keep").each do |keep_file|
          FileUtils.rm_rf(keep_file)
        end

        # Render/save generic stuff
        render_view('content_type.xml.erb', "#{extract_path}/[Content_Types].xml")
        render_view('presentation.xml.rel.erb', "#{extract_path}/ppt/_rels/presentation.xml.rels")
        render_view('presentation.xml.erb', "#{extract_path}/ppt/presentation.xml")
        render_view('app.xml.erb', "#{extract_path}/docProps/app.xml")

        # Save slides
        slides.each_with_index do |slide, index|
          slide.save(extract_path, index + 1)
        end

        # Create .pptx file
        File.delete(path) if File.exist?(path)
        Powerpoint.compress_pptx(extract_path, path)
      end

      path
    end

    def file_types
      types = slides.map do |slide|
        slide.file_types if slide.respond_to?(:file_types)
      end
      types.compact.flatten.uniq
    end
  end
end
