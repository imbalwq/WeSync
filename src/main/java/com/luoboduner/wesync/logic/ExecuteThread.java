package com.luoboduner.wesync.logic;

import com.luoboduner.wesync.App;
import com.luoboduner.wesync.ui.UiConsts;
import com.luoboduner.wesync.ui.panel.StatusPanel;
import com.luoboduner.wesync.vivo.LogReport;
import com.opencsv.CSVWriter;
import com.luoboduner.wesync.logic.bean.Table;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import com.luoboduner.wesync.tools.*;

import javax.swing.*;
import java.io.*;
import java.net.URL;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
            String tradosReportForder;
            String outputForder;

            tradosReportForder=StatusPanel.textFieldLogForder.getText().trim();
            outputForder=StatusPanel.textFieldLogForderOutput.getText().trim();
            boolean isUseFileNameSuffix=StatusPanel.buttonSelectValue==0?true:false;

            String msg="";
            //TODO 测试 start

            //TODO 读取模板文件并 写入到新的位置
            String templatePath= UiConsts.CURRENT_DIR + File.separator + "config" + File.separator + "reportTemplate.xlsx";
            File reportTemplate = new File(templatePath);

            Date now = new Date();

            String reportName="error.xlsx";
            if(StatusPanel.buttonSelectValue==0){
                reportName="vivo report.xlsx";
            }else if(StatusPanel.buttonSelectValue==1){
                reportName="log report.xlsx";
            }
            if(outputForder.trim().lastIndexOf("\\")==outputForder.trim().length()-1){
                outputForder=outputForder.substring(0,outputForder.length()-1)+"\\"+reportName;
            }else{
                outputForder=outputForder.substring(0,outputForder.length())+"\\"+reportName;
            }

//            LogReport.generateQuotationSheet(reportTemplate,reportFolder,newFolderPath,isUseFileNameSuffix);
            msg=LogReport.generateQuotationSheet(reportTemplate,tradosReportForder,outputForder,isUseFileNameSuffix);
            //TODO 测试 end
            StatusPanel.labelStatusDetail.setText("详情："+msg);
        } catch (Exception e) {
            isexception=true;
            logger.error("线程运行中异常：",e);
            JOptionPane.showMessageDialog(App.statusPanel, "报告汇总过程异常，可查[看详细日志]并联系工程部同事咨询！\n"+"错误信息概述："+e.getMessage());
            e.printStackTrace();
        }

        StatusPanel.buttonStartNow.setEnabled(true);
        StatusPanel.isRunning = false;
        if(!isexception){
            JOptionPane.showMessageDialog(App.statusPanel, "报告汇总完成！");
        }

        StatusPanel.progressCurrent.setValue(0);
    }
}
