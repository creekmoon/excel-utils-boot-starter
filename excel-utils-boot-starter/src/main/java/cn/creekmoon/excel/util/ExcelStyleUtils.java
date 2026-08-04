package cn.creekmoon.excel.util;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelStyleUtils  {



    public static void initDefaultStyle(Workbook workbook){
        /* POI 原生实现,等价于 Hutool StyleUtil.createDefaultCellStyle:水平垂直居中 + 四边细黑边框 */
        CellStyle defaultCellStyle = workbook.createCellStyle();
        defaultCellStyle.setAlignment(HorizontalAlignment.CENTER);
        defaultCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        defaultCellStyle.setBorderBottom(BorderStyle.THIN);
        defaultCellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        defaultCellStyle.setBorderLeft(BorderStyle.THIN);
        defaultCellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        defaultCellStyle.setBorderRight(BorderStyle.THIN);
        defaultCellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        defaultCellStyle.setBorderTop(BorderStyle.THIN);
        defaultCellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        defaultCellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        defaultCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    }


}
