require 'powerpoint'
require 'fastimage'

describe 'Powerpoint parsing a sample PPTX file' do
  before(:all) do 
    @deck = Powerpoint::Presentation.new
    @deck.add_intro 'Bicycle Of the Mind', 'created by Steve Jobs'
    @deck.add_textual_slide 'Why Mac?', ['Its cool!', 'Its light!']
    @deck.add_textual_slide 'Why Iphone?', ['Its fast!', 'Its cheap!']
    @deck.add_pictorial_slide 'JPG Logo', 'samples/images/sample_png.png'
    @deck.add_text_picture_slide('Text Pic Split', 'samples/images/sample_png.png', content = ['Here is a string', 'here is another'])
    @deck.add_pictorial_slide 'PNG Logo', 'samples/images/sample_png.png'
    @deck.add_picture_description_slide('Pic Desc', 'samples/images/sample_png.png', content = ['Here is a string', 'here is another'])
    @deck.add_picture_description_slide('JPG Logo', 'samples/images/sample_jpg.jpg', content = ['descriptions'])
    @deck.add_picture_description_slide('another', '/Users/stephenlink/Pictures/robothead.jpg', content = ['things'])
    @deck.add_pictorial_slide 'GIF Logo', 'samples/images/sample_gif.gif', {x: 124200, y: 3356451, cx: 2895600, cy: 1013460}
    @deck.add_textual_slide 'Why Android?', ['Its great!', 'Its sweet!']

    group_coords = { x: 838200, y: 1885950, cx: 7772400, cy: 4311650 }

    image_paths = ['samples/images/sample_png.png', 'samples/images/sample_png.png']
    image_sizes = image_paths.map { |path| FastImage.size(path) }

    coordinates = Powerpoint::Util.vertically_grouped_coordinates(group_coords, image_sizes)

    t = Powerpoint::Slide::Text

    elements = [
      Powerpoint::Slide::ImageContent.new(
        idx: 14,
        coords: coordinates[0],
        image_path: 'samples/images/sample_png.png'),

      Powerpoint::Slide::ImageContent.new(
        idx: 14,
        coords: coordinates[1],
        image_path: 'samples/images/sample_png.png'),

      Powerpoint::Slide::TextContent.new(
        idx: 13,
        content: ['long description line goes here']),

      Powerpoint::Slide::TextContent.new(
        idx: 16,
        content: [t.new('foo', size: 2400, color: '2d97d3') + 'bar', t.new('baz') + '◼']
      )
    ]

    slide_opts = {
        elements: elements,
        slide_layout: 12,
        title: 'foo bar slide'
    }

    @deck.add_slide(Powerpoint::Slide::DefaultSlide.new(slide_opts))



    @deck.save 'samples/pptx/sample.pptx' # Examine the PPTX file
  end

  it 'Create a PPTX file successfully.' do
    #@deck.should_not be_nil
  end
end
