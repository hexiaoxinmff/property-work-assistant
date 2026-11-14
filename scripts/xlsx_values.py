"""Read raw cells without mistaking Excel formulas for business values."""
def value_rows_without_formulas(sheet):
    for cells in sheet.iter_rows():
        for cell in cells:
            if cell.data_type == 'f':
                raise ValueError(f'{sheet.title}!{cell.coordinate} 为公式，须先核实计算结果及业务口径；不能将公式文本当作实际值')
        yield tuple(cell.value for cell in cells)
