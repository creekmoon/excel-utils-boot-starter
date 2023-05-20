package cn.creekmoon.excelUtils.core;

import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileNotFoundException;

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
        return getCustomAbsoluteFilePath(taskId + ".xlsx");
    }

    /**
     * 获取获取一个文件路径 默认路径在当前jar包的同级目录
     *
     * @param fileName 文件名称 需要自己保证不重复
     * @return
     */
    public static String getCustomAbsoluteFilePath(String fileName) {
//这里凌晨跨天的话会有问题 不过我不想管了... hhh
//        return getApplicationParentFilePath() + File.separator + "temp_files" + File.separator + DateUtil.format(new Date(), "yyyy-MM-dd") + File.separator + fileName;
        return getApplicationParentFilePath() + File.separator + "temp_files" + File.separator + fileName;
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
