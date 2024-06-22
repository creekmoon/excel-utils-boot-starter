package cn.creekmoon.excel.core.W;

import cn.creekmoon.excel.core.W.title.ext.ExcelCellStyle;
import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class Writer {

    @Getter
    protected Integer sheetIndex;
    @Getter
    protected String sheetName;

    abstract public Workbook getWorkbook();


    /**
     * 获取运行时的样式对象
     */
    abstract protected CellStyle getRunningTimeCellStyle(ExcelCellStyle style);


    /**
     * 声明周期钩子函数, 当写入数据时
     */
    protected void onWrite() {
    }

    ;

    /**
     * 声明周期钩子函数, 当切换sheet页时
     */
    protected void onSwitchSheet() {
    }

    ;

    /**
     * 声明周期钩子函数, 当执行stopWrite时
     */
    protected void onStopWrite() {
    }

    ;
}
