package com.xinyu.vivotool.logic;

import com.xinyu.vivotool.App;
import com.xinyu.vivotool.ui.UiConsts;
import com.xinyu.vivotool.ui.panel.StatusPanel;
import com.xinyu.vivotool.vivo.VivoExcelDataCopyMerge;
import com.xinyu.vivotool.vivo.VivoSplitExcel20231010;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


import javax.swing.*;
import java.io.*;
import java.util.Date;

/**
 * 执行器线程
 *
 * @author lwq
 */
public class ExecuteThread extends Thread  {

    private static final Logger logger = LoggerFactory.getLogger(ExecuteThread.class);

    /**
     * 初始化变量
     */
    public void init() {
//        tableFieldMap = new LinkedHashMap<>();
//        triggerMap = new LinkedHashMap<>();
//        originalTablesMap = new LinkedHashMap<>();
//        targetTablesMap = new LinkedHashMap<>();
    }

    @Override
    public void run() {
        StatusPanel.isRunning = true;
        this.setName("ExecuteThread");
        long enterTime = System.currentTimeMillis();
        // 初始化变量
        init();
        // 测试连接
        boolean isexception=false;
        try {
            //="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo LogReport\\log";
            //inputForder：译前文件夹 ， inputForder (译后目录，目录下需要包含批次文件夹 批次文件夹下需要包含 原文、译文)
            String inputForder;
            String outputForder;

            inputForder=StatusPanel.textFieldLogForder.getText().trim();
            outputForder=StatusPanel.textFieldLogForderOutput.getText().trim();
            boolean isUseFileNameSuffix=StatusPanel.buttonSelectValue==0?true:false;
            System.out.println("tradosReportForder:"+inputForder);
            System.out.println("outputForder"+outputForder);

            //TODO 读取模板文件并 写入到新的位置
            String templatePath= UiConsts.CURRENT_DIR + File.separator + "config" + File.separator + "vivoReportTemplate.xlsx";
            System.out.println("templatePath:"+templatePath);
            File reportTemplate = new File(templatePath);
            Date now = new Date();

            if(outputForder.trim().lastIndexOf("\\")==outputForder.trim().length()-1){
                outputForder=outputForder.substring(0,outputForder.length()-1)+"\\";
            }else{
                outputForder=outputForder.substring(0,outputForder.length())+"\\";
            }

            if(inputForder.trim().lastIndexOf("\\")==inputForder.trim().length()-1){
                inputForder=inputForder.substring(0,inputForder.length()-1)+"\\";
            }else{
                inputForder=inputForder.substring(0,inputForder.length())+"\\";
            }



//            LogReport.generateQuotationSheet(reportTemplate,reportFolder,newFolderPath,isUseFileNameSuffix);
//            msg=LogReport.generateQuotationSheet(reportTemplate,tradosReportForder,outputForder,isUseFileNameSuffix);
            String msg="";
            File outputDir = new File (outputForder);
            if(outputDir.exists() && outputDir.isDirectory()){
                //如果输出位置存在并且下边有其它文件，提示用户将先会清空目标位置
                File[] files = outputDir.listFiles();
                if(files != null && files.length > 0){
                    int result = JOptionPane.showConfirmDialog(App.statusPanel, "文件输出目录下存在其它文件，确认后将会先清空该目录是否是否继续？", "确认提示", JOptionPane.YES_NO_OPTION);
                    if(result == JOptionPane.YES_OPTION){
                        org.apache.commons.io.FileUtils.cleanDirectory(outputDir);
                    }else{
                        return;
                    }
                }
            }


            if(isUseFileNameSuffix){
            //TODO Vivo 译前处理
             msg= VivoSplitExcel20231010.vivoExcelSplit(inputForder,outputForder);
            }else{
                //TODO 译后处理
             msg= VivoExcelDataCopyMerge.vivoExcelDataMerge(inputForder,outputForder);
            }
            //TODO 测试 end
            StatusPanel.labelStatusDetail.setText("详情："+msg);
        } catch (Exception e) {
            isexception=true;
            logger.error("处理异常",e);
            e.printStackTrace();
        }

        StatusPanel.buttonStartNow.setEnabled(true);
        StatusPanel.isRunning = false;
        if(isexception){
            StatusPanel.labelStatusDetail.setText("处理失败，请查看详细日志！");
            JOptionPane.showMessageDialog(App.statusPanel, "处理失败，可查看详细日志！");
        }else{
            JOptionPane.showMessageDialog(App.statusPanel, "处理完成！");
        }

        StatusPanel.progressCurrent.setValue(0);
    }
}
