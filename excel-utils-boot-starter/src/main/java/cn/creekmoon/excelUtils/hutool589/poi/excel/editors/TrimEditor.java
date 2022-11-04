package cn.creekmoon.excelUtils.hutool589.poi.excel.editors;

import cn.creekmoon.excelUtils.hutool589.core.util.StrUtil;
import cn.creekmoon.excelUtils.hutool589.poi.excel.cell.CellEditor;
import org.apache.poi.ss.usermodel.Cell;

/**
 * 去除String类型的单元格值两边的空格
 *
 * @author Looly
 */
public class TrimEditor implements CellEditor {

	@Override
	public Object edit(Cell cell, Object value) {
		if (value instanceof String) {
			return StrUtil.trim((String) value);
		}
		return value;
	}

}
