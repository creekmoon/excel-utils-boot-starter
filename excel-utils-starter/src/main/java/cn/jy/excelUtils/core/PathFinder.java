package cn.jy.excelUtils.core;

import cn.hutool.core.date.DateUtil;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.Date;

/**
 * 获取路径
 */
public class PathFinder {

    /*临时文件目录*/
    private static volatile String applicationParentFilePath;

    /**
     * 获取文件绝对路径 默认路径在当前jar包的同级目录
     *
     * @param taskId 唯一名称,写文件完成后会获得
     * @return
     */
    public static String getAbsoluteFilePath(String taskId) {
        return getApplicationParentFilePath() + File.separator + "temp_files" + File.separator + DateUtil.format(new Date(), "yyyy-MM-dd") + File.separator + taskId + ".xlsx";
    }


    /*获取项目路径*/
    public static String getApplicationParentFilePath() {
        if (applicationParentFilePath == null) {
            synchronized (ExcelExport.class) {
                try {
                    if (applicationParentFilePath == null) {
                        applicationParentFilePath = new File(ResourceUtils.getURL("classpath:").getPath()).getParentFile().getParentFile().getParent();
                    }
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                }
            }
        }
        return applicationParentFilePath;
    }
}
