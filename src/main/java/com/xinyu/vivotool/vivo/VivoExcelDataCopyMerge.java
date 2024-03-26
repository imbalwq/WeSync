package com.xinyu.vivotool.vivo;

import javafx.beans.binding.StringBinding;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author liweiqing
 * @date 2023/7/26 9:28
 * @description
 */
public class VivoExcelDataCopyMerge {

    private static final Logger logger = LoggerFactory.getLogger(VivoSplitExcel.class);

    private static void main(String[] args) throws Exception {

//        String newFolderPath="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo result\\1、译前\\";

        String intputPath = "C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo result\\2、译后回填处理\\";
        String outputPath="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo result\\2、译文结果\\";
        vivoExcelDataMerge(intputPath, outputPath);
    }

    public static String vivoExcelDataMerge(String intputPath, String outputPath) throws Exception {
        //后边必须验证后加斜杠
        File[] file = (new File(intputPath)).listFiles();
        int errnum=0;
        Map<String, Integer> processingRecords = new HashMap<>();
        processingRecords.put("批次数量",0);
        processingRecords.put("原始文件数量",0);
        processingRecords.put("回稿文件数量",0);
        processingRecords.put("合并后件数量",0);
        for (File file1 : file) {
            boolean isSuccess=true;
            if(!file1.isFile()){
                String batchName=file1.getName();
                batchName=batchName.equals("")?"":batchName+"\\";
                logger.info("译后合并批次名："+batchName);
                //原文件目录 , 需要先手动解压文件, 压缩包暂时没有处理
                String originalFielForder= intputPath +batchName+"原文";
                //译文文件夹
                String translateFielForder= intputPath +batchName+"译文";
                String outputFolderPath= outputPath +batchName;
                //2、译后数据回填
                isSuccess=VivoSplitExcel.excelDataCopyMerge(originalFielForder,translateFielForder,outputFolderPath,processingRecords);
                if(!isSuccess){
                    errnum++;
                }
            }
        }
        if(errnum>0){
            return "译后处理完成，但其中有"+errnum+"个批次文件检测到异常，请查看【日志详情】！";
        }else{
            StringBuffer msg=new StringBuffer("");
            msg.append("译后处理完成， ")
                            .append("处理").append(processingRecords.get("批次数量")).append("个批次")
                            .append("，合计回稿数量：").append(processingRecords.get("回稿文件数量")).append("")
                            .append("，原文数量：").append(processingRecords.get("原始文件数量")).append("")
                            .append("，合并后文件数量：").append(processingRecords.get("合并后件数量")).append("。");
            return msg.toString();
        }
    }
}
