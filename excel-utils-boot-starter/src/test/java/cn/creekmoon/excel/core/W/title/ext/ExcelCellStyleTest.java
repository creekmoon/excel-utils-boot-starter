package cn.creekmoon.excel.core.W.title.ext;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.util.function.BiConsumer;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * ExcelCellStyle 预设与工具方法契约：
 * 所有预设应可应用到 POI CellStyle 而不抛异常，merge/rgb/hex/font 行为符合声明。
 */
class ExcelCellStyleTest {

    private static final BiConsumer<Workbook, CellStyle>[] ALL_PRESETS = new BiConsumer[]{
            ExcelCellStyle.LIGHT_ORANGE.getStyleInitializer(),
            ExcelCellStyle.PALE_BLUE.getStyleInitializer(),
            ExcelCellStyle.LIGHT_GREEN.getStyleInitializer(),
            ExcelCellStyle.HEADER_PRIMARY.getStyleInitializer(),
            ExcelCellStyle.HEADER_SUCCESS.getStyleInitializer(),
            ExcelCellStyle.HEADER_WARNING.getStyleInitializer(),
            ExcelCellStyle.HEADER_DANGER.getStyleInitializer(),
            ExcelCellStyle.HEADER_INFO.getStyleInitializer(),
            ExcelCellStyle.HEADER_DARK.getStyleInitializer(),
            ExcelCellStyle.STATUS_SUCCESS.getStyleInitializer(),
            ExcelCellStyle.STATUS_WARNING.getStyleInitializer(),
            ExcelCellStyle.STATUS_DANGER.getStyleInitializer(),
            ExcelCellStyle.STATUS_INFO.getStyleInitializer(),
            ExcelCellStyle.STRIPE_EVEN.getStyleInitializer(),
            ExcelCellStyle.BORDER_THIN_ALL.getStyleInitializer(),
            ExcelCellStyle.BORDER_BOTTOM_BOLD.getStyleInitializer(),
            ExcelCellStyle.ALIGN_CENTER.getStyleInitializer(),
            ExcelCellStyle.ALIGN_RIGHT.getStyleInitializer(),
            ExcelCellStyle.ALIGN_LEFT.getStyleInitializer(),
            ExcelCellStyle.FORMAT_CURRENCY.getStyleInitializer(),
            ExcelCellStyle.FORMAT_PERCENTAGE.getStyleInitializer(),
            ExcelCellStyle.FORMAT_INTEGER.getStyleInitializer(),
            ExcelCellStyle.FORMAT_DATE.getStyleInitializer(),
            ExcelCellStyle.FORMAT_DATETIME.getStyleInitializer(),
            ExcelCellStyle.FONT_BOLD.getStyleInitializer(),
            ExcelCellStyle.FONT_ITALIC.getStyleInitializer(),
            ExcelCellStyle.HEADER_PRIMARY_BORDER.getStyleInitializer(),
            ExcelCellStyle.STATUS_SUCCESS_BORDER.getStyleInitializer(),
            ExcelCellStyle.STATUS_WARNING_BORDER.getStyleInitializer(),
            ExcelCellStyle.STATUS_DANGER_BORDER.getStyleInitializer(),
            ExcelCellStyle.STATUS_INFO_BORDER.getStyleInitializer(),
    };

    @Test
    void allPresets_applyToWorkbookStyle_withoutError() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            for (BiConsumer<Workbook, CellStyle> preset : ALL_PRESETS) {
                CellStyle style = workbook.createCellStyle();
                preset.accept(workbook, style);
                assertNotNull(style);
            }
        }
    }

    @Test
    void headerPrimary_setsBoldWhiteCenteredStyle() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            CellStyle style = workbook.createCellStyle();
            ExcelCellStyle.HEADER_PRIMARY.getStyleInitializer().accept(workbook, style);
            assertTrue(workbook.getFontAt(style.getFontIndex()).getBold());
            assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
            assertEquals(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER, style.getAlignment());
        }
    }

    @Test
    void merge_appliesStylesInOrder_secondOverridesFirst() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            CellStyle style = workbook.createCellStyle();
            // 合并：橙色填充 + 全边框
            ExcelCellStyle.merge(ExcelCellStyle.LIGHT_ORANGE, ExcelCellStyle.BORDER_THIN_ALL)
                    .getStyleInitializer().accept(workbook, style);
            assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
            assertEquals(org.apache.poi.ss.usermodel.BorderStyle.THIN, style.getBorderTop());
            assertEquals(org.apache.poi.ss.usermodel.BorderStyle.THIN, style.getBorderBottom());
        }
    }

    @Test
    void rgb_createsExactForegroundColor() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            CellStyle style = workbook.createCellStyle();
            ExcelCellStyle.rgb(0x1A, 0x5F, 0x7A).getStyleInitializer().accept(workbook, style);
            assertTrue(style instanceof XSSFCellStyle);
            XSSFColor color = ((XSSFCellStyle) style).getFillForegroundXSSFColor();
            assertArrayEquals(new byte[]{(byte) 0x1A, (byte) 0x5F, (byte) 0x7A}, color.getRGB());
        }
    }

    @Test
    void hex_decomposesIntoSameRgb() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            CellStyle style = workbook.createCellStyle();
            ExcelCellStyle.hex(0x3498DB).getStyleInitializer().accept(workbook, style);
            XSSFColor color = ((XSSFCellStyle) style).getFillForegroundXSSFColor();
            assertArrayEquals(new byte[]{(byte) 0x34, (byte) 0x98, (byte) 0xDB}, color.getRGB());
        }
    }

    @Test
    void fontConfig_buildsConfiguredFont() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Font font = new ExcelCellStyle.FontConfig().bold().italic().size(14).build(workbook);
            assertTrue(font.getBold());
            assertTrue(font.getItalic());
            assertEquals(14, font.getFontHeightInPoints());
        }
    }

    @Test
    void currencyFormat_setsDataFormat() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            CellStyle style = workbook.createCellStyle();
            ExcelCellStyle.FORMAT_CURRENCY.getStyleInitializer().accept(workbook, style);
            assertTrue(style.getDataFormat() != 0, "currency style must set a data format");
        }
    }

    @Test
    void conditionCellStyle_holdsStyleAndPredicate() {
        ExcelCellStyle style = ExcelCellStyle.LIGHT_GREEN;
        ConditionCellStyle<Integer> condition = ConditionCellStyle.of(style, x -> x > 5);
        assertEquals(style, condition.getDefaultCellStyle());
        assertTrue(condition.getCondition().test(10));
        assertTrue(!condition.getCondition().test(1));
    }
}
