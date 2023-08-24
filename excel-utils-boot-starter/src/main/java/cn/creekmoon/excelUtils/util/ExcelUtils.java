package cn.creekmoon.excelUtils.util;

public class ExcelUtils {

    /**
     * excel列号转数字
     * A-->1   AA-->27
     *
     * @param column 列号字母 例如A,B,C
     * @return
     */
    public static int excelColumnToNumber(String column) {
        int result = 0;
        for (int i = 0; i < column.length(); i++) {
            char c = column.charAt(i);
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }


    /**
     * excel单元格获取列号数字
     * 例如输入 B6  返回2    输入AA23 返回27
     *
     * @param
     * @return
     */
    public static int excelCellToColumnNumber(String cell) {
        String column = cell.replaceAll("[^A-Za-z]", "");
        return excelColumnToNumber(column);
    }

    /**
     * excel单元格获取行号数字
     * 例如输入 B6  返回6    输入AA23 返回23
     *
     * @param
     * @return
     */
    public static int excelCellToRowNumber(String cell) {
        String row = cell.replaceAll("[^0-9]", "");
        return Integer.parseInt(row);
    }


    public static void main(String[] args) {
        System.out.println("excelCellToRowNumber(\"B28\") = " + excelCellToRowNumber("B28"));
    }
}
