package cn.creekmoon.excel.core.W.title.ext;

import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.function.BiConsumer;
import java.util.function.Consumer;

public class ExcelCellStyle {

    public final static ExcelCellStyle LIGHT_ORANGE = new ExcelCellStyle(cellStyle ->
    {
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });


    public final static ExcelCellStyle PALE_BLUE = new ExcelCellStyle(cellStyle ->
    {
        cellStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });


    public final static ExcelCellStyle LIGHT_GREEN = new ExcelCellStyle(cellStyle ->
    {
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });


    @Getter
    private BiConsumer<Workbook, CellStyle> styleInitializer;

    public ExcelCellStyle(Consumer<CellStyle> styleInitializer) {
        this.styleInitializer = (x, y) -> styleInitializer.accept(y);
    }


    public ExcelCellStyle(BiConsumer<Workbook, CellStyle> styleInitializer) {
        this.styleInitializer = styleInitializer;
    }
}