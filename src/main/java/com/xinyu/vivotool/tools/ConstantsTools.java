package com.xinyu.vivotool.tools;

import com.xinyu.vivotool.ui.UiConsts;

import javax.swing.*;
import java.io.File;

/**
 * 工具层相关的常量
 *
 * @author lwq
 */
public class ConstantsTools {

    // 配置文件
    /**
     * 配置文件 路径
     */
    public final static String PATH_CONFIG = UiConsts.CURRENT_DIR + File.separator + "config" + File.separator
            + "config.xml";

    /**
     * properties路径
     */
    public final static String PATH_PROPERTY = UiConsts.CURRENT_DIR + File.separator + "config" + File.separator
            + "zh-cn.properties";
    /**
     * 配置文件dom实例
     */
    public final static ConfigManager CONFIGER = ConfigManager.getConfigManager();
    /**
     * xpath
     */
    public final static String XPATH_AUTO_BAK = "//vivo/setting/autoBak";
    public final static String XPATH_MYSQL_PATH = "//vivo/setting/mysqlPath";

    public final static String XPATH_CHECKBOX_SPLIT_INPUT_PATH = "//vivo/setting/cbxSplitDataInput";
    public final static String XPATH_TEXT_SPLIT_INPUT_PATH = "//vivo/setting/splitDataInputForder";
    public final static String XPATH_CHECKBOX_SPLIT_OUTPUT_PATH = "//vivo/setting/cbxSplitDataOutput";
    public final static String XPATH_TEXT_SPLIT_OUTPUT_PATH = "//vivo/setting/splitDataOutputForder";
    public final static String XPATH_CHECKBOX_MERGE_INPUT_PATH = "//vivo/setting/cbxMergeDataInput";
    public final static String XPATH_TEXT_MERGE_INPUT_PATH = "//vivo/setting/mergeDataInputForder";
    public final static String XPATH_CHECKBOX_MERGE_OUTPUT_PATH = "//vivo/setting/cbxMergeDataOutput";
    public final static String XPATH_TEXT_MERGE_OUTPUT_PATH = "//vivo/setting/mergeDataOutputForder";


    public final static String XPATH_DEBUG_MODE = "//weSync/setting/debugMode";
    /**
     * 日志文件 路径
     */
    public final static String PATH_LOG = UiConsts.CURRENT_DIR + File.separator + "log" + File.separator + "log.log";
}
