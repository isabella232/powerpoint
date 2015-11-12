require 'powerpoint'

describe 'Content' do
  it 'generates image xml' do

    img_opts = {
        id: 1,
        idx: 4,
        embed_rid: 'rId2',
        coords: {x: 1, cx: 10, y: 2, cy: 20},
        image_path: 'samples/images/sample_png.png'
    }

    expect {
      str = Powerpoint::Slide::ImageContent.new(img_opts).render
    }.not_to raise_error
  end

  it 'generates table xml' do

    table_opts = {
        id: 1,
        idx: 4,
        content: [%w(a b c), %w(1 2 3)]
    }

    expect {
      str = Powerpoint::Slide::TableContent.new(table_opts).render
    }.not_to raise_error

  end

  it 'generates text xml' do

    text_opts = {
        id: 1,
        idx: 4,
        content: %w(line another third)
    }

    expect {
      str = Powerpoint::Slide::TextContent.new(text_opts).render
    }.not_to raise_error
  end

  it 'default slide has content' do

    table_opts = {
        id: 11,
        idx: 14,
        content: [%w(a b c), %w(1 2 3)]
    }

    text_opts = {
        id: 10,
        idx: 13,
        content: %w(line another third)
    }

    elements = [
        Powerpoint::Slide::TableContent.new(table_opts),
        Powerpoint::Slide::TextContent.new(text_opts)
    ]

    slide_opts = {
        elements: elements,
        slide_layout: 12,
        title: 'foo bar slide'
    }

    slide = Powerpoint::Slide::DefaultSlide.new(slide_opts)
  end

  it 'strings text together' do
    t = Powerpoint::Slide::Text

    str = t.new('foo', size: 1200, color: 'f12g59') + t.new(' bar') + t.new(' baz')
    rendered = str.render

    expect(rendered).to include('foo')
    expect(rendered).to include('f12g59')
    expect(rendered).to include('sz="1200"')
    expect(rendered).to include('bar')
    expect(rendered).to include('baz')
  end

end






