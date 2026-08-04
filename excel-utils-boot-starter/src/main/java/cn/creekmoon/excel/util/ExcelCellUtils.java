package cn.creekmoon.excel.util;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelCellUtils {

    /**
     * excel列号转数字(0-based)
     * A-->0   AA-->26
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
     * excel单元格获取列号数字(0-based)
     * 例如输入 B6  返回1    输入AA23 返回26
     *
     * @param
     * @return
     */
    public static int excelCellToColumnIndex(String cell) {
        String column = cell.replaceAll("[^A-Za-z]", "");
        return excelColumnToIndex(column);
    }

    /**
     * excel单元格获取行号数字(0-based)
     * 例如输入 B6  返回5    输入AA23 返回22
     *
     * @param
     * @return
     */
    public static int excelCellToRowIndex(String cell) {
        String row = cell.replaceAll("[^0-9]", "");
        return Integer.parseInt(row) - 1;
    }

    /**
     * excel行号(0-based)和列号(0-based)转为1-based单元格引用
     * 例如输入 0,0 返回A1
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

        // POI 原生写文件示例(替代 Hutool ExcelWriter)
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream out = new FileOutputStream("C:\\Users\\JY\\Desktop\\草稿.xlsx")) {
            workbook.createSheet().createRow(0).createCell(0).setCellValue("A15555555555555555");
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
