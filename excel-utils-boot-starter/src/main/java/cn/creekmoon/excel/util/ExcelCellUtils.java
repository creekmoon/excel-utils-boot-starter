package cn.creekmoon.excel.util;

import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;

public class ExcelCellUtils {

    /**
     * excel列号转数字
     * A-->1   AA-->27
     *
     * @param column 列号字母 例如A,B,C
     * @return
     */
    public static int excelColumnToIndex(String column) {
        int result = 0;
        for (int i = 0; i < column.length(); i++) {
            char c = column.charAt(i);
            result = result * 26 + (c - 'A' + 1);
        }
        return result - 1;
    }


    /**
     * excel单元格获取列号数字
     * 例如输入 B6  返回2    输入AA23 返回27
     *
     * @param
     * @return
     */
    public static int excelCellToColumnIndex(String cell) {
        String column = cell.replaceAll("[^A-Za-z]", "");
        return excelColumnToIndex(column);
    }

    /**
     * excel单元格获取行号数字
     * 例如输入 B6  返回6    输入AA23 返回23
     *
     * @param
     * @return
     */
    public static int excelCellToRowIndex(String cell) {
        String row = cell.replaceAll("[^0-9]", "");
        return Integer.parseInt(row) - 1;
    }

    /**
     * excel列号和行号转为CellReference
     * 例如输入 6,2 返回B6
     *
     * @param
     * @return
     */
    public static String excelIndexToCell(Integer rowIndex, Integer colIndex) {

        int col = colIndex + 1;
        int row = rowIndex + 1;
        StringBuilder sb = new StringBuilder();

        while (col > 0) {
            int remainder = (col - 1) % 26;
            char c = (char) ('A' + remainder);
            sb.insert(0, c);
            col = (col - 1) / 26;
        }

        sb.append(row);

        return sb.toString();
    }

    public static void main(String[] args) {
        System.out.println("excelCellToRowNumber(\"B28\") = " + excelCellToRowIndex("B28"));
        System.out.println("excelIndexToCell(2,1) = " + excelIndexToCell(27, 26));

        ExcelWriter writer = ExcelUtil.getWriter("C:\\Users\\JY\\Desktop\\草稿.xlsx");
        writer.writeCellValue(0, 0, "A15555555555555555");
        writer.flush();
        writer.close();
    }
}
