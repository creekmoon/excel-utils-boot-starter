package cn.creekmoon.excelUtils.hutool589.poi.excel;

import cn.creekmoon.excelUtils.hutool589.core.io.IORuntimeException;
import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

/**
 * Excel文件工具类
 *
 * @author looly
 * @since 4.2.1
 */
public class ExcelFileUtil {
    // ------------------------------------------------------------------------------------------------ isXls

    /**
     * 是否为XLS格式的Excel文件（HSSF）<br>
     * XLS文件主要用于Excel 97~2003创建<br>
     * 此方法会自动调用{@link InputStream#reset()}方法
     *
     * @param in excel输入流
     * @return 是否为XLS格式的Excel文件（HSSF）
     */
    public static boolean isXls(InputStream in) {
        return FileMagic.OLE2 == getFileMagic(in);
    }

    /**
     * 是否为XLSX格式的Excel文件（XSSF）<br>
     * XLSX文件主要用于Excel 2007+创建<br>
     * 此方法会自动调用{@link InputStream#reset()}方法
     *
     * @param in excel输入流
     * @return 是否为XLSX格式的Excel文件（XSSF）
     */
    public static boolean isXlsx(InputStream in) {
        return FileMagic.OOXML == getFileMagic(in);
    }

    /**
     * 是否为XLSX格式的Excel文件（XSSF）<br>
     * XLSX文件主要用于Excel 2007+创建
     *
     * @param file excel文件
     * @return 是否为XLSX格式的Excel文件（XSSF）
     * @since 5.4.4
     */
    public static boolean isXlsx(File file) {
        try {
            return FileMagic.valueOf(file) == FileMagic.OOXML;
        } catch (IOException e) {
            throw new IORuntimeException(e);
        }
    }

    /**
     * {@link java.io.PushbackInputStream}
     * PushbackInputStream的markSupported()为false，并不支持mark和reset
     * 如果强转成PushbackInputStream在调用FileMagic.valueOf(inputStream)时会报错
     * {@link FileMagic}
     * 报错内容：getFileMagic() only operates on streams which support mark(int)
     * 此处修改成 final InputStream in = FileMagic.prepareToCheckMagic(in)
     *
     * @param in {@link InputStream}
     * @author kefan.qu
     */
    private static FileMagic getFileMagic(InputStream in) {
        FileMagic magic;
        in = FileMagic.prepareToCheckMagic(in);
        try {
            magic = FileMagic.valueOf(in);
        } catch (IOException e) {
            throw new IORuntimeException(e);
        }

        return magic;
    }
}
