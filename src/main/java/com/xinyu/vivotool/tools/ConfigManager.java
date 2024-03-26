package com.xinyu.vivotool.tools;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.io.OutputFormat;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 配置文件管理 单例
 *
 * @author lwq
 */
public class ConfigManager {
    private volatile static ConfigManager confManager;
    private static final Logger logger = LoggerFactory.getLogger(ConfigManager.class);
    public Document document;

    /**
     * 私有的构造
     */
    private ConfigManager() {
        reloadDom();
    }

    /**
     * 读取xml，加载到dom
     */
    public void reloadDom() {
        SAXReader reader = new SAXReader();
        try {
            document = reader.read(new File(ConstantsTools.PATH_CONFIG));
        } catch (DocumentException e) {
            e.printStackTrace();
            logger.error("Read config xml error:" + e.toString());
        }
    }

    /**
     * 获取实例，线程安全
     *
     * @return
     */
    public static ConfigManager getConfigManager() {
        if (confManager == null) {
            synchronized (ConfigManager.class) {
                if (confManager == null) {
                    confManager = new ConfigManager();
                }
            }
        }
        return confManager;
    }

    /**
     * 把document对象写入新的文件
     *
     * @throws Exception
     */
    public void writeToXml() throws Exception {
        // 排版缩进的格式
        OutputFormat format = OutputFormat.createPrettyPrint();
        // 设置编码
        format.setEncoding("UTF-8");
        // 创建XMLWriter对象,指定了写出文件及编码格式
        XMLWriter writer = null;
        writer = new XMLWriter(
                new OutputStreamWriter(new FileOutputStream(new File(ConstantsTools.PATH_CONFIG)), StandardCharsets.UTF_8), format);

        // 写入
        writer.write(document);
        writer.flush();
        writer.close();

    }


    public String getAutoBak() {
        return this.document.selectSingleNode(ConstantsTools.XPATH_AUTO_BAK).getText();
    }

    public void setAutoBak(String autoBak) throws Exception {
        this.document.selectSingleNode(ConstantsTools.XPATH_AUTO_BAK).setText(autoBak);
        writeToXml();
    }

    public String getMysqlPath() {
        return this.document.selectSingleNode(ConstantsTools.XPATH_MYSQL_PATH).getText();
    }

    public void setMysqlPath(String mysqlPath) throws Exception {
        this.document.selectSingleNode(ConstantsTools.XPATH_MYSQL_PATH).setText(mysqlPath);
        writeToXml();
    }


    /**vivo 译前拆分 原文输入 */
    public String getXpathCheckboxSplitInputPath() {
        return this.document.selectSingleNode(ConstantsTools.XPATH_CHECKBOX_SPLIT_INPUT_PATH).getText();
    }
    public void setXpathCheckboxSplitInputPath(String param) throws Exception {
        this.document.selectSingleNode(ConstantsTools.XPATH_CHECKBOX_SPLIT_INPUT_PATH).setText(param);
        writeToXml();
    }

    public String getXpathTextSplitInputPath() {
        return this.document.selectSingleNode(ConstantsTools.XPATH_TEXT_SPLIT_INPUT_PATH).getText();
    }
    public void setXpathTextSplitInputPath(String param) throws Exception {
        this.document.selectSingleNode(ConstantsTools.XPATH_TEXT_SPLIT_INPUT_PATH).setText(param);
        writeToXml();
    }

    /**vivo 译前拆分 输出 */
    public String getXpathCheckboxSplitOutputPath() {
        return this.document.selectSingleNode(ConstantsTools.XPATH_CHECKBOX_SPLIT_OUTPUT_PATH).getText();
    }
    public void setXpathCheckboxSplitOutputPath(String param) throws Exception {
        this.document.selectSingleNode(ConstantsTools.XPATH_CHECKBOX_SPLIT_OUTPUT_PATH).setText(param);
        writeToXml();
    }

    public String getXpathTextSplitOutputPath() {
        return this.document.selectSingleNode(ConstantsTools.XPATH_TEXT_SPLIT_OUTPUT_PATH).getText();
    }
    public void setXpathTextSplitOutputPath(String param) throws Exception {
        this.document.selectSingleNode(ConstantsTools.XPATH_TEXT_SPLIT_OUTPUT_PATH).setText(param);
        writeToXml();
    }


    /** 译后输入输出 */
    public String getXpathCheckboxMergeInputPath() {
        return this.document.selectSingleNode(ConstantsTools.XPATH_CHECKBOX_MERGE_INPUT_PATH).getText();
    }
    public void setXpathCheckboxMergeInputPath(String param) throws Exception {
        this.document.selectSingleNode(ConstantsTools.XPATH_CHECKBOX_MERGE_INPUT_PATH).setText(param);
        writeToXml();
    }

    public String getXpathTextMergeInputPath() {
        return this.document.selectSingleNode(ConstantsTools.XPATH_TEXT_MERGE_INPUT_PATH).getText();
    }
    public void setXpathTextMergeInputPath(String param) throws Exception {
        this.document.selectSingleNode(ConstantsTools.XPATH_TEXT_MERGE_INPUT_PATH).setText(param);
        writeToXml();
    }

    public String getXpathCheckboxMergeOutputPath() {
        return this.document.selectSingleNode(ConstantsTools.XPATH_CHECKBOX_MERGE_OUTPUT_PATH).getText();
    }
    public void setXpathCheckboxMergeOutputPath(String param) throws Exception {
        this.document.selectSingleNode(ConstantsTools.XPATH_CHECKBOX_MERGE_OUTPUT_PATH).setText(param);
        writeToXml();
    }

    public String getXpathTextMergeOutputPath() {
        return this.document.selectSingleNode(ConstantsTools.XPATH_TEXT_MERGE_OUTPUT_PATH).getText();
    }
    public void setXpathTextMergeOutputPath(String param) throws Exception {
        this.document.selectSingleNode(ConstantsTools.XPATH_TEXT_MERGE_OUTPUT_PATH).setText(param);
        writeToXml();
    }

    public String getDebugMode() {
        return this.document.selectSingleNode(ConstantsTools.XPATH_DEBUG_MODE).getText();
    }

    public void setDebugMode(String debugMode) throws Exception {
        this.document.selectSingleNode(ConstantsTools.XPATH_DEBUG_MODE).setText(debugMode);
        writeToXml();
    }
}
