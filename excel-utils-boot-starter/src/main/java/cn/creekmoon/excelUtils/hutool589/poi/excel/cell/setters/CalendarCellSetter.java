package cn.creekmoon.excelUtils.hutool589.poi.excel.cell.setters;

import cn.creekmoon.excelUtils.hutool589.poi.excel.cell.CellSetter;
import org.apache.poi.ss.usermodel.Cell;

import java.util.Calendar;

/**
 * {@link Calendar} 值单元格设置器
 *
 * @author looly
 * @since 5.7.8
 */
public class CalendarCellSetter implements CellSetter {

    private final Calendar value;

    /**
     * 构造
     *
     * @param value 值
     */
    CalendarCellSetter(Calendar value) {
        this.value = value;
    }

    @Override
    public void setValue(Cell cell) {
        cell.setCellValue(value);
    }
}
