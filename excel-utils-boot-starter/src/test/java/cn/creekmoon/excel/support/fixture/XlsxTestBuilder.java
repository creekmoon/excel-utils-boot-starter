package cn.creekmoon.excel.support.fixture;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.mock.web.MockMultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.Date;
import java.util.List;

/**
 * 测试用动态 xlsx 生成器：按二维行数据在内存中生成 xlsx 字节，
 * 用于构造空白行、乱序列、多余列、日期单元格等静态 fixture 不便覆盖的场景。
 *
 * <p>单元格类型规则：String→文本，Integer/Long→数字，Double/BigDecimal→数字，Date/LocalDateTime→日期单元格，null→空白。</p>
 * <p>空白"行"请传入只含空字符串的 Object[]（SAX 才会读到该行，可用于验证空白行过滤）。</p>
 */
public final class XlsxTestBuilder {

    private XlsxTestBuilder() {
    }

    /** 生成单 sheet xlsx 字节（sheet 名 Sheet0）。 */
    public static byte[] xlsxBytes(List<Object[]> rows) throws IOException {
        return xlsxBytes("Sheet0", rows);
    }

    /** 生成单 sheet xlsx 字节。 */
    public static byte[] xlsxBytes(String sheetName, List<Object[]> rows) throws IOException {
        return xlsxBytes(new String[]{sheetName}, List.of(rows));
    }

    /** 生成多 sheet xlsx 字节，sheetNames 与 sheets 按索引一一对应。 */
    public static byte[] xlsxBytes(String[] sheetNames, List<List<Object[]>> sheets) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            for (int i = 0; i < sheetNames.length; i++) {
                Sheet sheet = workbook.createSheet(sheetNames[i]);
                List<Object[]> rows = sheets.get(i);
                for (int r = 0; r < rows.size(); r++) {
                    Row row = sheet.createRow(r);
                    Object[] values = rows.get(r);
                    for (int c = 0; c < values.length; c++) {
                        writeCell(row.createCell(c), values[c]);
                    }
                }
            }
            workbook.write(out);
            return out.toByteArray();
        }
    }

    /** 生成单 sheet MockMultipartFile。 */
    public static MockMultipartFile multipartFile(String fileName, List<Object[]> rows) throws IOException {
        return new MockMultipartFile(fileName, fileName, null, xlsxBytes(rows));
    }

    private static void writeCell(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof String s) {
            cell.setCellValue(s);
        } else if (value instanceof Integer i) {
            cell.setCellValue(i);
        } else if (value instanceof Long l) {
            cell.setCellValue(l);
        } else if (value instanceof Double d) {
            cell.setCellValue(d);
        } else if (value instanceof Float f) {
            cell.setCellValue(f);
        } else if (value instanceof java.math.BigDecimal bd) {
            cell.setCellValue(bd.doubleValue());
        } else if (value instanceof Boolean b) {
            cell.setCellValue(b);
        } else if (value instanceof Date date) {
            cell.setCellValue(date);
        } else if (value instanceof LocalDateTime ldt) {
            cell.setCellValue(Date.from(ldt.atZone(ZoneId.systemDefault()).toInstant()));
        } else {
            cell.setCellValue(String.valueOf(value));
        }
    }
}
