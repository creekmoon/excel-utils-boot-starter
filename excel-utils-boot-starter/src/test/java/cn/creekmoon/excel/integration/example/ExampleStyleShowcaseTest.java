package cn.creekmoon.excel.integration.example;

import cn.creekmoon.excel.core.W.ExcelExport;
import cn.creekmoon.excel.support.assertion.PoiWorkbookAssertions;
import cn.creekmoon.excel.support.fixture.StyleDemo;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.List;

import static cn.creekmoon.excel.core.W.title.ext.ConditionCellStyle.of;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.ALIGN_CENTER;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.ALIGN_LEFT;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.ALIGN_RIGHT;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.BORDER_THIN_ALL;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.FONT_BOLD;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.FONT_ITALIC;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.FORMAT_CURRENCY;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.FORMAT_DATE;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.FORMAT_DATETIME;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.FORMAT_INTEGER;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.FORMAT_PERCENTAGE;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.HEADER_PRIMARY;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.STATUS_DANGER;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.STATUS_DANGER_BORDER;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.STATUS_SUCCESS;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.STATUS_SUCCESS_BORDER;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.STATUS_WARNING;
import static cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.STATUS_WARNING_BORDER;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * 复刻 excel-utils-example /export/style-showcase：8 个 sheet 的样式预设集成验证。
 * 单元层已验证预设可应用，此处验证「真实导出文件」中样式确实落盘（颜色/格式/字体/对齐）。
 */
class ExampleStyleShowcaseTest {

    private static final int DATA_ROWS = 4;

    @Test
    void styleShowcase_allEightSheets_stylesAppliedCorrectly() throws Exception {
        List<StyleDemo> data = StyleDemo.createShowcaseData();

        ExcelExport excelExport = ExcelExport.create();

        /* ===== Sheet 1: 表头样式（基础导出） ===== */
        excelExport.switchNewSheet("1-表头样式", StyleDemo.class)
                .addTitle("姓名", StyleDemo::getName)
                .addTitle("年龄", StyleDemo::getAge)
                .addTitle("分数", StyleDemo::getScore)
                .write(data);

        /* ===== Sheet 2: 状态色（条件样式） ===== */
        excelExport.switchNewSheet("2-状态色", StyleDemo.class)
                .addTitle("姓名", StyleDemo::getName)
                .addTitle("状态", StyleDemo::getStatus,
                        of(STATUS_SUCCESS, d -> "正常".equals(d.getStatus())),
                        of(STATUS_WARNING, d -> "警告".equals(d.getStatus())),
                        of(STATUS_DANGER, d -> "异常".equals(d.getStatus())))
                .addTitle("分数", StyleDemo::getScore,
                        of(STATUS_SUCCESS_BORDER, d -> d.getScore() >= 80),
                        of(STATUS_WARNING_BORDER, d -> d.getScore() >= 60 && d.getScore() < 80),
                        of(STATUS_DANGER_BORDER, d -> d.getScore() < 60))
                .write(data);

        /* ===== Sheet 3: 斑马纹 + 边框 ===== */
        excelExport.switchNewSheet("3-斑马纹+边框", StyleDemo.class)
                .addTitle("姓名", StyleDemo::getName, BORDER_THIN_ALL, d -> true)
                .addTitle("年龄", StyleDemo::getAge, BORDER_THIN_ALL, d -> true)
                .addTitle("类别", StyleDemo::getCategory, BORDER_THIN_ALL, d -> true)
                .addTitle("分数", StyleDemo::getScore, BORDER_THIN_ALL, d -> true)
                .write(data);

        /* ===== Sheet 4: 数字格式 ===== */
        excelExport.switchNewSheet("4-数字格式", StyleDemo.class)
                .addTitle("姓名", StyleDemo::getName)
                .addTitle("货币", StyleDemo::getAmount, FORMAT_CURRENCY, d -> true)
                .addTitle("百分比", StyleDemo::getRate, FORMAT_PERCENTAGE, d -> true)
                .addTitle("千分位", StyleDemo::getScore, FORMAT_INTEGER, d -> true)
                .addTitle("日期", StyleDemo::getDate, FORMAT_DATE, d -> true)
                .addTitle("日期时间", StyleDemo::getDatetime, FORMAT_DATETIME, d -> true)
                .write(data);

        /* ===== Sheet 5: 对齐样式 ===== */
        excelExport.switchNewSheet("5-对齐样式", StyleDemo.class)
                .addTitle("左对齐", StyleDemo::getName, ALIGN_LEFT, d -> true)
                .addTitle("居中", StyleDemo::getStatus, ALIGN_CENTER, d -> true)
                .addTitle("右对齐", StyleDemo::getScore, ALIGN_RIGHT, d -> true)
                .write(data);

        /* ===== Sheet 6: 组合样式 ===== */
        excelExport.switchNewSheet("6-组合样式", StyleDemo.class)
                .addTitle("表头+边框+居中", StyleDemo::getName,
                        cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.merge(HEADER_PRIMARY, BORDER_THIN_ALL, ALIGN_CENTER), d -> true)
                .addTitle("状态+边框+右对齐", StyleDemo::getScore,
                        of(cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.merge(STATUS_SUCCESS, BORDER_THIN_ALL, ALIGN_RIGHT),
                                d -> d.getScore() >= 80))
                .write(data);

        /* ===== Sheet 7: 自定义颜色 ===== */
        excelExport.switchNewSheet("7-自定义颜色", StyleDemo.class)
                .addTitle("HEX 色值", StyleDemo::getName,
                        of(cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.hex(0xE8D5B7), d -> true))
                .addTitle("RGB 色值", StyleDemo::getStatus,
                        of(cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.rgb(180, 220, 250), d -> true))
                .addTitle("IndexedColor", StyleDemo::getCategory,
                        of(cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.hex(0xDDA0DD), d -> true))
                .write(data);

        /* ===== Sheet 8: 字体样式 ===== */
        excelExport.switchNewSheet("8-字体样式", StyleDemo.class)
                .addTitle("粗体", StyleDemo::getName, FONT_BOLD, d -> true)
                .addTitle("斜体", StyleDemo::getStatus, FONT_ITALIC, d -> true)
                .addTitle("粗体+居中", StyleDemo::getScore,
                        cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle.merge(FONT_BOLD, ALIGN_CENTER), d -> true)
                .write(data);

        File file = excelExport.stopWrite();
        assertTrue(file.exists());

        try (Workbook workbook = PoiWorkbookAssertions.open(file)) {
            PoiWorkbookAssertions.assertSheetNames(workbook,
                    "1-表头样式", "2-状态色", "3-斑马纹+边框", "4-数字格式",
                    "5-对齐样式", "6-组合样式", "7-自定义颜色", "8-字体样式");

            /* Sheet2 状态色：张三(正常)→SUCCESS、李四(警告)→WARNING、王五(异常)→DANGER */
            Sheet sheet2 = workbook.getSheetAt(1);
            assertFillRgb(sheet2, 1, 1, 0xD5, 0xF5, 0xE3);   // STATUS_SUCCESS
            assertFillRgb(sheet2, 2, 1, 0xFC, 0xF3, 0xCF);   // STATUS_WARNING
            assertFillRgb(sheet2, 3, 1, 0xFA, 0xDB, 0xD8);   // STATUS_DANGER
            /* 分数列：85→SUCCESS_BORDER、72→WARNING_BORDER、58→DANGER_BORDER、91→SUCCESS_BORDER */
            assertFillRgb(sheet2, 1, 2, 0xD5, 0xF5, 0xE3);
            assertFillRgb(sheet2, 2, 2, 0xFC, 0xF3, 0xCF);
            assertFillRgb(sheet2, 3, 2, 0xFA, 0xDB, 0xD8);
            assertFillRgb(sheet2, 4, 2, 0xD5, 0xF5, 0xE3);
            assertEquals(org.apache.poi.ss.usermodel.BorderStyle.THIN,
                    sheet2.getRow(1).getCell(2).getCellStyle().getBorderTop(), "SUCCESS_BORDER 应含边框");

            /* Sheet3 全边框 */
            Sheet sheet3 = workbook.getSheetAt(2);
            assertEquals(org.apache.poi.ss.usermodel.BorderStyle.THIN,
                    sheet3.getRow(1).getCell(0).getCellStyle().getBorderTop());

            /* Sheet4 数字格式 */
            Sheet sheet4 = workbook.getSheetAt(3);
            assertDataFormat(sheet4, 1, 1, "#,##0.00");                    // 货币 ¥#,##0.00
            assertDataFormat(sheet4, 1, 2, "0.00%");                       // 百分比
            assertDataFormat(sheet4, 1, 3, "#,##0");                       // 千分位
            assertDataFormat(sheet4, 1, 4, "yyyy-MM-dd");                  // 日期
            assertDataFormat(sheet4, 1, 5, "yyyy-MM-dd HH:mm:ss");         // 日期时间

            /* Sheet5 对齐 */
            Sheet sheet5 = workbook.getSheetAt(4);
            assertEquals(HorizontalAlignment.LEFT, sheet5.getRow(1).getCell(0).getCellStyle().getAlignment());
            assertEquals(HorizontalAlignment.CENTER, sheet5.getRow(1).getCell(1).getCellStyle().getAlignment());
            assertEquals(HorizontalAlignment.RIGHT, sheet5.getRow(1).getCell(2).getCellStyle().getAlignment());

            /* Sheet6 组合：表头色 + 边框 + 居中 */
            Sheet sheet6 = workbook.getSheetAt(5);
            CellStyle merged = sheet6.getRow(1).getCell(0).getCellStyle();
            assertFillRgb(sheet6, 1, 0, 0x1A, 0x5F, 0x7A);                 // HEADER_PRIMARY 深青
            assertEquals(org.apache.poi.ss.usermodel.BorderStyle.THIN, merged.getBorderTop());
            assertEquals(HorizontalAlignment.CENTER, merged.getAlignment());

            /* Sheet7 自定义颜色：hex 与 rgb 精确色值 */
            Sheet sheet7 = workbook.getSheetAt(6);
            assertFillRgb(sheet7, 1, 0, 0xE8, 0xD5, 0xB7);                 // hex(0xE8D5B7)
            assertFillRgb(sheet7, 1, 1, 0xB4, 0xDC, 0xFA);                 // rgb(180,220,250)
            assertFillRgb(sheet7, 1, 2, 0xDD, 0xA0, 0xDD);                 // hex(0xDDA0DD)

            /* Sheet8 字体：粗体 / 斜体 / 粗体+居中 */
            Sheet sheet8 = workbook.getSheetAt(7);
            assertTrue(workbook.getFontAt(sheet8.getRow(1).getCell(0).getCellStyle().getFontIndex()).getBold());
            assertTrue(workbook.getFontAt(sheet8.getRow(1).getCell(1).getCellStyle().getFontIndex()).getItalic());
            CellStyle boldCenter = sheet8.getRow(1).getCell(2).getCellStyle();
            assertTrue(workbook.getFontAt(boldCenter.getFontIndex()).getBold());
            assertEquals(HorizontalAlignment.CENTER, boldCenter.getAlignment());

            /* 所有 sheet 行数 = 表头 1 + 数据 4 */
            for (int i = 0; i < 8; i++) {
                PoiWorkbookAssertions.assertRowCount(workbook.getSheetAt(i), DATA_ROWS + 1);
            }
        }
    }

    private static void assertFillRgb(Sheet sheet, int row, int col, int r, int g, int b) {
        Cell cell = sheet.getRow(row).getCell(col);
        CellStyle style = cell.getCellStyle();
        assertTrue(style instanceof XSSFCellStyle, "XSSFWorkbook 下应为 XSSFCellStyle");
        XSSFColor color = ((XSSFCellStyle) style).getFillForegroundXSSFColor();
        assertArrayEquals(new byte[]{(byte) r, (byte) g, (byte) b}, color.getRGB(),
                "单元格 (" + row + "," + col + ") 填充色");
    }

    private static void assertDataFormat(Sheet sheet, int row, int col, String format) {
        CellStyle style = sheet.getRow(row).getCell(col).getCellStyle();
        String actual = style.getDataFormatString();
        assertTrue(actual != null && actual.contains(format),
                "单元格 (" + row + "," + col + ") 数字格式应为 " + format + "，实际: " + actual);
    }
}
