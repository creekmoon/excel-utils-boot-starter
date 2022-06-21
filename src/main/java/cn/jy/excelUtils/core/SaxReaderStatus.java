package cn.jy.excelUtils.core;

import lombok.Data;

import java.util.HashMap;
import java.util.Map;


/**
 * sax模式读取状态
 */
@Data
public class SaxReaderStatus {

    /*成功读取并消费的行号*/
    int successRowIndex = 0;
    /*key=行号 value=错误信息*/
    Map<Integer, String> errorReport = new HashMap<>();
}
