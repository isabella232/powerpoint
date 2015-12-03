require 'powerpoint'

describe 'Cell' do
  it 'calculates size' do


    top_left = Powerpoint::Slide::TableCell.new('top')

    cell2 = Powerpoint::Slide::TableCell.new('cell2')
    cell3 = Powerpoint::Slide::TableCell.new('cell3')
    cell4 = Powerpoint::Slide::TableCell.new('cell4')


    top_left.set_right_neighbors([cell2, cell3])

    cell2.set_bottom_neighbors([cell3, cell4])

    cell3.set_right_neighbors(cell4)


    h, v = top_left.calculate_sizing
    expect(h).to eq(1)
    expect(v).to eq(2)

    h, v = cell2.calculate_sizing
    expect(h).to eq(2)
    expect(v).to eq(1)

  end

  it 'finds rightmost and bottommost neighbors' do
    cell = Powerpoint::Slide::TableCell


    top_left = cell.new('top')

    cell2 = cell.new('cell2')
    cell3 = cell.new('cell3')
    cell4 = cell.new('cell4')


    top_left.set_right_neighbors([cell2, cell3])

    cell2.set_bottom_neighbors([cell3, cell4])

    cell3.set_right_neighbors(cell4)


    expect(top_left.right_most).to eq([cell2, cell4])

    expect(top_left.top_most).to eq([top_left, cell2])
    expect(top_left.bottom_most).to eq([top_left, cell3, cell4])

    expect(cell2.bottom_most).to eq([cell3, cell4])

    expect(cell2.left_most).to eq([cell2, cell3])
    expect(top_left.left_most).to eq([top_left])

  end

  it 'welds things correctly' do
    cell = Powerpoint::Slide::TableCell


    def verifyCells(cells, expected_data)
      expect(cells.map { |r| r.data } ).to eq(expected_data)
    end

    top_left = cell.new('topleft')
    bottom_left = cell.row(['botleft1', 'botleft2', 'botleft3'])

    verifyCells(bottom_left.right_most, ['botleft3'])
    verifyCells(bottom_left.collect_rightwards_stay_down, ['botleft1', 'botleft2', 'botleft3'])

    left_side = cell.weld_vertical(top_left, bottom_left)

    verifyCells(top_left.bottom_most, ['botleft1', 'botleft2', 'botleft3'])

    right_side = cell.column(['topright', 'botright'])

    out = cell.weld_horizontal(left_side, right_side)

    verifyCells(out.right_most, ['topright', 'botright'])
    verifyCells(out.bottom_most, ['botleft1', 'botleft2', 'botleft3', 'botright'])

    row2 = cell.row(['realbotleft', 'p1', 'p2', 'realbotright'])

    verifyCells(row2.right_most, ['realbotright'])

    out2 = cell.weld_vertical(out, row2)

    verifyCells(row2.right_most, ['realbotright'])
    verifyCells(out2.bottom_most, ['realbotleft', 'p1', 'p2', 'realbotright'])
    verifyCells(out2.bottom_most, ['realbotleft', 'p1', 'p2', 'realbotright'])

    h, v = out2.calculate_sizing

    expect(h).to eq(3)
    expect(v).to eq(1)

  end

  it 'can be assembled into a renderable square' do
    cell = Powerpoint::Slide::TableCell

    top_left = cell.new('topleft')
    bottom_left = cell.row(['botleft1', 'botleft2', 'botleft3'])
    left_side = cell.weld_vertical(top_left, bottom_left)


    # not welding so we can get a square arrangement here
    right_col = cell.column(['topright1', 'topright2', 'botright'])
    top_left.set_right_neighbors([right_col, right_col.bottom[0]])
    bottom_left.all_right[0].set_right_neighbors(right_col.bottom_left)

    h, v = top_left.calculate_sizing

    expect(h).to eq(3)
    expect(v).to eq(2)

    renderable = top_left.to_renderable

    expect(renderable.size).to eq(3)
    renderable.each do |row|
      expect(row.size).to eq(4)
    end


    Powerpoint::Slide::EnhancedTableContent.new(id: 1, idx: 1, table_cells: top_left).render

  end
end
