package cn.creekmoon.excel.util;

import cn.hutool.poi.excel.style.StyleUtil;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelStyleUtils  {



    public static void initDefaultStyle(Workbook workbook){
        CellStyle defaultCellStyle = StyleUtil.createDefaultCellStyle(workbook);
        defaultCellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        defaultCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

    }


}
