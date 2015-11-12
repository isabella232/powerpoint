require 'powerpoint'

describe 'Content' do

  it 'converts pixel to pt' do
    pt = Powerpoint::Util.pixle_to_pt(1)
    expect(pt).to eq(12700)
  end

  it 'scales image height to fit' do
    image_coords = { x: 100*12700, y: 200*12700, cx: 100*12700, cy: 200*12700 }
    dimensions = [50, 50] # square should get scaled to fit x but not y

    output_coords = Powerpoint::Util.boxed_coordinates(image_coords, dimensions)
    expect(output_coords).to eq({ x: 100*12700, y: 200*12700, cx: 100*12700, cy: 100*12700})
  end

  it 'scales image width to fit' do
    image_coords = { x: 100*12700, y: 200*12700, cx: 100*12700, cy: 200*12700 }
    dimensions = [10, 50] # should get scaled to fit y but not x

    output_coords = Powerpoint::Util.boxed_coordinates(image_coords, dimensions)
    expect(output_coords).to eq({ x: 100*12700, y: 200*12700, cx: 40*12700, cy: 200*12700})
  end

  it 'scales down to fit the smaller dimension' do
    image_coords = { x: 100*12700, y: 200*12700, cx: 100*12700, cy: 200*12700 }
    dimensions = [50*10000000, 50*10000000] # should get scaled to fit y but not x

    output_coords = Powerpoint::Util.boxed_coordinates(image_coords, dimensions)
    expect(output_coords).to eq({ x: 100*12700, y: 200*12700, cx: 100*12700, cy: 100*12700})
  end

  it 'scales an image group to fit with max width' do
    template_coords = { x: 100*12700, y: 200*12700, cx: 100*12700, cy: 200*12700 }
    dimensions = [
      [50, 50], # should get scaled to fit y but not x
      [75, 25],
      [60, 25],
    ]

    output_coords = Powerpoint::Util.vertically_grouped_coordinates(template_coords, dimensions)

    total_y = output_coords.map { |c| c[:cy]}.reduce(:+)
    max_x = output_coords.map { |c| c[:cx]}.max

    expect(total_y).to be <= template_coords[:cy]
    expect(max_x).to eq(template_coords[:cx])
  end

  it 'scales an image group to fit with max length' do
    template_coords = { x: 100*12700, y: 200*12700, cx: 100*12700, cy: 200*12700 }
    dimensions = [
        [10, 50], # should get scaled to fit y but not x
        [15, 25],
        [10, 25],
    ]

    output_coords = Powerpoint::Util.vertically_grouped_coordinates(template_coords, dimensions)

    total_y = output_coords.map { |c| c[:cy]}.reduce(:+)
    max_x = output_coords.map { |c| c[:cx]}.max

    expect(total_y).to eq(template_coords[:cy])
    expect(max_x).to be <= template_coords[:cx]
  end
end
