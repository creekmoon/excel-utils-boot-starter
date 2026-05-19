package cn.creekmoon.excel.core.W.title.ext;

import lombok.Getter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.util.function.BiConsumer;
import java.util.function.Consumer;

/**
 * Excel 单元格样式定义。
 *
 * <p>提供两类预设：</p>
 * <ul>
 *   <li><b>常量预设</b>：直接引用即可，如 {@link #HEADER_PRIMARY}、{@link #STATUS_SUCCESS}</li>
 *   <li><b>工具方法</b>：{@link #rgb(int, int, int)}、{@link #font(Workbook, FontConfig)} 等，用于自定义颜色/字体</li>
 * </ul>
 *
 * <p>所有预设均基于 POI 5.x + .xlsx（XSSF）格式设计。</p>
 */
public class ExcelCellStyle {

    /* ======================== 兼容旧版（保留） ======================== */

    public final static ExcelCellStyle LIGHT_ORANGE = new ExcelCellStyle(cellStyle -> {
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });

    public final static ExcelCellStyle PALE_BLUE = new ExcelCellStyle(cellStyle -> {
        cellStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });

    public final static ExcelCellStyle LIGHT_GREEN = new ExcelCellStyle(cellStyle -> {
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });

    /* ======================== 表头样式 ======================== */

    /** 主色表头：深青底 + 白字 + 居中 + 粗体 */
    public final static ExcelCellStyle HEADER_PRIMARY = new ExcelCellStyle((workbook, style) -> {
        fillRgb(style, 0x1A, 0x5F, 0x7A);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = font(workbook, f -> {
            f.setBold(true);
            f.setColor(IndexedColors.WHITE.getIndex());
            f.setFontHeightInPoints((short) 11);
        });
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    });

    /** 成功表头：翠绿底 + 白字 + 居中 + 粗体 */
    public final static ExcelCellStyle HEADER_SUCCESS = new ExcelCellStyle((workbook, style) -> {
        fillRgb(style, 0x27, 0xAE, 0x60);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = font(workbook, f -> {
            f.setBold(true);
            f.setColor(IndexedColors.WHITE.getIndex());
            f.setFontHeightInPoints((short) 11);
        });
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    });

    /** 警告表头：琥珀底 + 白字 + 居中 + 粗体 */
    public final static ExcelCellStyle HEADER_WARNING = new ExcelCellStyle((workbook, style) -> {
        fillRgb(style, 0xF3, 0x9C, 0x12);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = font(workbook, f -> {
            f.setBold(true);
            f.setColor(IndexedColors.WHITE.getIndex());
            f.setFontHeightInPoints((short) 11);
        });
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    });

    /** 危险表头：砖红底 + 白字 + 居中 + 粗体 */
    public final static ExcelCellStyle HEADER_DANGER = new ExcelCellStyle((workbook, style) -> {
        fillRgb(style, 0xE7, 0x4C, 0x3C);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = font(workbook, f -> {
            f.setBold(true);
            f.setColor(IndexedColors.WHITE.getIndex());
            f.setFontHeightInPoints((short) 11);
        });
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    });

    /** 信息表头：亮蓝底 + 白字 + 居中 + 粗体 */
    public final static ExcelCellStyle HEADER_INFO = new ExcelCellStyle((workbook, style) -> {
        fillRgb(style, 0x34, 0x98, 0xDB);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = font(workbook, f -> {
            f.setBold(true);
            f.setColor(IndexedColors.WHITE.getIndex());
            f.setFontHeightInPoints((short) 11);
        });
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    });

    /** 深色表头：深灰蓝底 + 白字 + 居中 + 粗体 */
    public final static ExcelCellStyle HEADER_DARK = new ExcelCellStyle((workbook, style) -> {
        fillRgb(style, 0x2C, 0x3E, 0x50);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = font(workbook, f -> {
            f.setBold(true);
            f.setColor(IndexedColors.WHITE.getIndex());
            f.setFontHeightInPoints((short) 11);
        });
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    });

    /* ======================== 状态样式（数据行） ======================== */

    /** 成功状态：薄荷绿背景 */
    public final static ExcelCellStyle STATUS_SUCCESS = new ExcelCellStyle(cellStyle -> {
        fillRgb(cellStyle, 0xD5, 0xF5, 0xE3);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });

    /** 警告状态：奶油黄背景 */
    public final static ExcelCellStyle STATUS_WARNING = new ExcelCellStyle(cellStyle -> {
        fillRgb(cellStyle, 0xFC, 0xF3, 0xCF);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });

    /** 危险状态：淡珊瑚背景 */
    public final static ExcelCellStyle STATUS_DANGER = new ExcelCellStyle(cellStyle -> {
        fillRgb(cellStyle, 0xFA, 0xDB, 0xD8);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });

    /** 信息状态：冰蓝背景 */
    public final static ExcelCellStyle STATUS_INFO = new ExcelCellStyle(cellStyle -> {
        fillRgb(cellStyle, 0xD6, 0xEA, 0xF8);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });

    /* ======================== 斑马纹 ======================== */

    /** 偶数行斑马纹：极浅灰背景 */
    public final static ExcelCellStyle STRIPE_EVEN = new ExcelCellStyle(cellStyle -> {
        fillRgb(cellStyle, 0xF8, 0xF9, 0xFA);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });

    /* ======================== 边框样式 ======================== */

    /** 全细边框 + 浅灰色 */
    public final static ExcelCellStyle BORDER_THIN_ALL = new ExcelCellStyle(cellStyle -> {
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        cellStyle.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        cellStyle.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        cellStyle.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
    });

    /** 底部加粗边框 */
    public final static ExcelCellStyle BORDER_BOTTOM_BOLD = new ExcelCellStyle(cellStyle -> {
        cellStyle.setBorderBottom(BorderStyle.MEDIUM);
        cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
    });

    /* ======================== 对齐样式 ======================== */

    /** 水平居中 */
    public final static ExcelCellStyle ALIGN_CENTER = new ExcelCellStyle(cellStyle -> {
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    });

    /** 水平右对齐 */
    public final static ExcelCellStyle ALIGN_RIGHT = new ExcelCellStyle(cellStyle -> {
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    });

    /** 水平左对齐 */
    public final static ExcelCellStyle ALIGN_LEFT = new ExcelCellStyle(cellStyle -> {
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
    });

    /* ======================== 数字格式 ======================== */

    /** 货币格式：¥#,##0.00 */
    public final static ExcelCellStyle FORMAT_CURRENCY = new ExcelCellStyle((workbook, style) -> {
        style.setDataFormat(workbook.createDataFormat().getFormat("¥#,##0.00"));
    });

    /** 百分比格式：0.00% */
    public final static ExcelCellStyle FORMAT_PERCENTAGE = new ExcelCellStyle((workbook, style) -> {
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
    });

    /** 千分位整数：#,##0 */
    public final static ExcelCellStyle FORMAT_INTEGER = new ExcelCellStyle((workbook, style) -> {
        style.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
    });

    /** 日期格式：yyyy-MM-dd */
    public final static ExcelCellStyle FORMAT_DATE = new ExcelCellStyle((workbook, style) -> {
        style.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd"));
    });

    /** 日期时间格式：yyyy-MM-dd HH:mm:ss */
    public final static ExcelCellStyle FORMAT_DATETIME = new ExcelCellStyle((workbook, style) -> {
        style.setDataFormat(workbook.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss"));
    });

    /* ======================== 字体样式 ======================== */

    /** 粗体 */
    public final static ExcelCellStyle FONT_BOLD = new ExcelCellStyle((workbook, style) -> {
        Font font = font(workbook, f -> f.setBold(true));
        style.setFont(font);
    });

    /** 斜体 */
    public final static ExcelCellStyle FONT_ITALIC = new ExcelCellStyle((workbook, style) -> {
        Font font = font(workbook, f -> f.setItalic(true));
        style.setFont(font);
    });

    /* ======================== 组合快捷样式 ======================== */

    /** 表头 + 全边框 */
    public final static ExcelCellStyle HEADER_PRIMARY_BORDER = merge(HEADER_PRIMARY, BORDER_THIN_ALL);

    /** 成功状态 + 全边框 */
    public final static ExcelCellStyle STATUS_SUCCESS_BORDER = merge(STATUS_SUCCESS, BORDER_THIN_ALL);

    /** 警告状态 + 全边框 */
    public final static ExcelCellStyle STATUS_WARNING_BORDER = merge(STATUS_WARNING, BORDER_THIN_ALL);

    /** 危险状态 + 全边框 */
    public final static ExcelCellStyle STATUS_DANGER_BORDER = merge(STATUS_DANGER, BORDER_THIN_ALL);

    /** 信息状态 + 全边框 */
    public final static ExcelCellStyle STATUS_INFO_BORDER = merge(STATUS_INFO, BORDER_THIN_ALL);


    /* ======================== 核心属性 ======================== */

    @Getter
    private BiConsumer<Workbook, CellStyle> styleInitializer;

    public ExcelCellStyle(Consumer<CellStyle> styleInitializer) {
        this.styleInitializer = (x, y) -> styleInitializer.accept(y);
    }

    public ExcelCellStyle(BiConsumer<Workbook, CellStyle> styleInitializer) {
        this.styleInitializer = styleInitializer;
    }


    /* ======================== 工具方法 ======================== */

    /**
     * 将多个 ExcelCellStyle 合并为一个（后覆盖前）。
     *
     * <p>适用场景：表头样式 + 边框样式、状态样式 + 对齐样式 等组合。</p>
     *
     * @param styles 要合并的样式列表
     * @return 合并后的新样式
     */
    public static ExcelCellStyle merge(ExcelCellStyle... styles) {
        return new ExcelCellStyle((workbook, cellStyle) -> {
            for (ExcelCellStyle s : styles) {
                if (s != null && s.styleInitializer != null) {
                    s.styleInitializer.accept(workbook, cellStyle);
                }
            }
        });
    }

    /**
     * 用 RGB 值创建纯色背景样式。
     *
     * @param r 红 0-255
     * @param g 绿 0-255
     * @param b 蓝 0-255
     * @return 纯色背景样式
     */
    public static ExcelCellStyle rgb(int r, int g, int b) {
        return new ExcelCellStyle(cellStyle -> fillRgb(cellStyle, r, g, b));
    }

    /**
     * 用 16 进制色值创建纯色背景样式。
     *
     * @param hex 如 0x3498DB
     * @return 纯色背景样式
     */
    public static ExcelCellStyle hex(int hex) {
        int r = (hex >> 16) & 0xFF;
        int g = (hex >> 8) & 0xFF;
        int b = hex & 0xFF;
        return rgb(r, g, b);
    }

    /**
     * 为 CellStyle 设置 RGB 填充色（内部使用 XSSFColor）。
     * <p>注意：仅适用于 .xlsx 格式（XSSFCellStyle）。</p>
     */
    public static void fillRgb(CellStyle cellStyle, int r, int g, int b) {
        if (cellStyle instanceof XSSFCellStyle xssfStyle) {
            XSSFColor color = new XSSFColor(new byte[]{(byte) r, (byte) g, (byte) b}, null);
            xssfStyle.setFillForegroundColor(color);
        } else {
            // 降级：使用 IndexedColors.WHITE（HSSF 不支持自定义 RGB）
            cellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        }
    }

    /**
     * 创建并配置 Font 的便捷方法。
     *
     * @param workbook 工作簿
     * @param config   字体配置 Lambda
     * @return 配置好的 Font
     */
    public static Font font(Workbook workbook, Consumer<Font> config) {
        Font font = workbook.createFont();
        if (config != null) {
            config.accept(font);
        }
        return font;
    }

    /**
     * 字体配置 DSL 辅助类，便于创建复杂字体样式。
     */
    public static class FontConfig {
        public boolean bold = false;
        public boolean italic = false;
        public short color = IndexedColors.BLACK.getIndex();
        public short fontHeightInPoints = 11;

        public FontConfig bold() {
            this.bold = true;
            return this;
        }

        public FontConfig italic() {
            this.italic = true;
            return this;
        }

        public FontConfig color(short indexedColor) {
            this.color = indexedColor;
            return this;
        }

        public FontConfig color(int r, int g, int b) {
            // XSSF 字体颜色需要通过 XSSFFont.setColor(XSSFColor) 设置，这里暂存索引色
            this.color = IndexedColors.BLACK.getIndex();
            return this;
        }

        public FontConfig size(int pt) {
            this.fontHeightInPoints = (short) pt;
            return this;
        }

        public Font build(Workbook workbook) {
            Font font = workbook.createFont();
            font.setBold(bold);
            font.setItalic(italic);
            font.setColor(color);
            font.setFontHeightInPoints(fontHeightInPoints);
            return font;
        }
    }
}
