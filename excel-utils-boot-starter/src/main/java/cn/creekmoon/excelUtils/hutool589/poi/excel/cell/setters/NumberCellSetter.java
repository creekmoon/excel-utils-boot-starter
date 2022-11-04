package cn.creekmoon.excelUtils.hutool589.poi.excel.cell.setters;

import cn.creekmoon.excelUtils.hutool589.core.util.NumberUtil;
import cn.creekmoon.excelUtils.hutool589.poi.excel.cell.CellSetter;
import org.apache.poi.ss.usermodel.Cell;

/**
 * {@link Number} 值单元格设置器
 *
 * @author looly
 * @since 5.7.8
 */
public class NumberCellSetter implements CellSetter {

    private final Number value;

    /**
     * 构造
     *
     * @param value 值
     */
    NumberCellSetter(Number value) {
        this.value = value;
    }

    @Override
    public void setValue(Cell cell) {
        // issue https://gitee.com/dromara/hutool/issues/I43U9G
        // 避免float到double的精度问题
        cell.setCellValue(NumberUtil.toDouble(value));
    }
}
