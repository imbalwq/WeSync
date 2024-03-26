package com.xinyu.vivotool.vivo;


import com.xinyu.vivotool.vivo.util.ExcelColumns;
import com.xinyu.vivotool.vivo.util.FileUtil;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * @author liweiqing
 * @date 2023/4/10 14:52
 * @description
 */
public class VivoSplitExcel {

    private static FormulaEvaluator evaluator;
    private static String [] ExcelColumnsMark;
    private static final Logger logger = LoggerFactory.getLogger(VivoSplitExcel.class);

    /**
     *
     * @param originalFielForder   原始文件目录 客户原始文件
     * @param translateFielForder  翻译后的返回的excel文件目录
     * @param outputFolderPath     将合并后的数据输出到的目录
     */
    public static boolean excelDataCopyMerge(String originalFielForder, String translateFielForder,String outputFolderPath,Map<String,Integer> processingRecords) throws IOException {

//        originalFielForder="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo result\\2、译后回填处理\\原文";
//        translateFielForder="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo result\\2、译后回填处理\\译文";
//        outputFolderPath="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo result\\2、译后回填处理\\2、译后回填完成文件";

        boolean isSuccess=true;
        //TODO 初始化 Excel 表头与索引关系对应
        ExcelColumnsMark = ExcelColumns.getColumnLabels(200);

        //清空结果文件夹目录
        File deldir1 = new File (outputFolderPath);
        deldir1.mkdirs();
        FileUtils.cleanDirectory(deldir1);

        File destDir = new File (outputFolderPath);
        if(!(destDir.exists() && destDir.isDirectory())){
            destDir.mkdirs();
        }

        //原始文件集合
        List<File> listOriginalFiles = FileUtil.readyAllFiel(originalFielForder,"xlsx");
        //译后回稿文件集合
        List<File> listTranslateFiels= FileUtil.readyAllFiel(translateFielForder);
        //译文回稿文件 以文件名为 key 转换为 map
        Map<String, File> trFilesMap = listTranslateFiels.stream().collect(Collectors.toMap(File::getName, file -> file));

        //原始文件 map 以文件名为 key
        Map<String, File> originalFileMap = new HashMap<>();
        //需要 copy的原始文件集合 ，因为同一个批次 译文可能分几次返回
        Map<String, File> newFileMap = new HashMap<>();
        //原始文件存入map,用于后续查看哪些需要 copy
        for (File file : listOriginalFiles) {
            Map<String, Integer> map=null;
            if (file.isFile()) {
                String fileName=file.getName();
                String fileSuffix=fileName.substring(fileName.lastIndexOf("."));
                if(fileSuffix.equalsIgnoreCase(".xlsx")){
                    // Excel 数据拆分处理
                    originalFileMap.put(file.getName(), file);
                }
            }
        }
        //根据返回的文件 查看需要将哪些原文件 copy到新目录用于后续的数据合并使用
        for (File file : listTranslateFiels) {
            Map<String, Integer> map=null;
            if (file.isFile()) {
                String fileName=file.getName();
                String fileSuffix=fileName.substring(fileName.lastIndexOf("."));
                if(fileSuffix.equalsIgnoreCase(".xlsx")){
                    String delTargetLanguageCodeFileName=fileName.substring(0,fileName.lastIndexOf("-"))+".xlsx";
                    File originalFile=originalFileMap.get(delTargetLanguageCodeFileName);
                    if(originalFileMap.get(delTargetLanguageCodeFileName)!=null){
                        if(newFileMap.get(delTargetLanguageCodeFileName)==null){

                            // Excel 数据拆分处理
                            File newFile=new File(outputFolderPath+"\\"+delTargetLanguageCodeFileName);
                            //TODO 需要先删除该目录
                            Files.copy(originalFile.toPath(), newFile.toPath());
                            newFileMap.put(delTargetLanguageCodeFileName,newFile);
                        }
                    }else{
                        logger.error("回稿文件 无法匹配到原文件做数据回填：");
                        logger.error("--------------------回稿文件名："+fileName);
                        logger.error("--------------------应匹配的原文件名为："+delTargetLanguageCodeFileName);
                        isSuccess=false;
                    }
                }
            }
        }

        //TODO 将回稿文件的数据 回填到 copy出来的原始文件中。
        //读取新文件检查需要回填的语言 有哪些，所有目标语言的回稿文件是否都存在
        for (Map.Entry<String, File> newfile : newFileMap.entrySet()) {
            File file=newfile.getValue();
            FileInputStream excelFileInputStreamTemp = new FileInputStream(file.getPath());
            XSSFWorkbook workbook = null;
            try{
                ZipSecureFile.setMinInflateRatio(-1.0d);
                workbook = new XSSFWorkbook(excelFileInputStreamTemp);
            }catch (Exception e){
                logger.error(file.getName(),e);
                e.printStackTrace();
                return false;
            }

            excelFileInputStreamTemp.close();
            XSSFSheet sheet = workbook.getSheetAt(0);
            //TODO 读取部分数据 用于判断 该文件应如何归类 及 目标语言表头的定位 end
            List<String> targetLanguageList = new ArrayList<>();
            targetLanguageList.add("vnl");
            targetLanguageList.add("vhr_rHR");
            targetLanguageList.add("vfi_rFI");
            targetLanguageList.add("vta_rIN");
            targetLanguageList.add("vkk_rKZ");
            targetLanguageList.add("ves_rUS");
            //2023年9月01日新增语言 start
            targetLanguageList.add("vzh_rHK");
            targetLanguageList.add("vzh_rTW");
            targetLanguageList.add("vtl");
            targetLanguageList.add("var_rEG");
            targetLanguageList.add("vbn_rBD");
            targetLanguageList.add("vpt_rPT");
            targetLanguageList.add("ves_rES");
            targetLanguageList.add("vhi");
            targetLanguageList.add("vca_rES");
            targetLanguageList.add("vuk_rUA");
            targetLanguageList.add("vsr_rRS");
            targetLanguageList.add("vfa");
            targetLanguageList.add("vka");
            targetLanguageList.add("vmn_rMN");
            targetLanguageList.add("vpt_rBR");
            targetLanguageList.add("ven_rGB");
            targetLanguageList.add("vha");
            //2023年9月01日新增语言 end
            targetLanguageList.add("vmy_rMM");
            targetLanguageList.add("vmy_rZG");
            targetLanguageList.add("vms");
            targetLanguageList.add("vtr_rTR");

            //记录该文件中 目标语言的索引位置
            Map<String,Integer> targetLanguageIndexMap = new HashMap<>();
            //记录原始文件中，原文(简中、英文)索引位置
            Map<String, Integer> sourceLanguageIndexMap = new HashMap<>();
            //TODO 检查文件名 根据文件名，判断属于 新增翻译 还是 核对文件
            boolean isNewTranslate=true;
            //应用类型 在D列， 表头为："应用类型" 暂时固定
//            int applicationTypeIndex=3;
            String newFolderPath2="";
            if(file.getName().toLowerCase().contains("retranslate")){
                //TODO execuetReTranslate()  业务暂时合并没做拆分
                isNewTranslate=false;
            }else if(file.getName().toLowerCase().contains("newtranslate")){
                //TODO execuetNewTranslate()    业务暂时合并没做拆分
                isNewTranslate=true;
            }else{
                logger.error("无法通过文件名确定是新增翻译或核对， 文件名为："+file.getName());
                logger.error("路径："+file.getPath());
                return false;
            }
            //遍历表头确定语言码位置及原语言表头位置
            Row row0=sheet.getRow(0);
            // 理论上表头不会存在重复的目标语言
            for(int columnIndex=0;columnIndex<row0.getLastCellNum();columnIndex++){
                String title=getCellValueByCell(row0.getCell(columnIndex)).trim();

                for (String s : targetLanguageList) {
                    if(title.equalsIgnoreCase(s)){
                        targetLanguageIndexMap.put(title,columnIndex);
                    }
                }
                if("vzh_rCN".equalsIgnoreCase(title) || "vus".equalsIgnoreCase(title) || "应用类型".equalsIgnoreCase(title)){
                    sourceLanguageIndexMap.put(title,columnIndex);
                }
            }

            //根据原文，检查回稿文件是否完整，  targetLanguageIndex:原文件 目标语言的索引位置
            out1:for (Map.Entry<String, Integer> targetLanguageIndex : targetLanguageIndexMap.entrySet()) {
                String key=file.getName().substring(0,file.getName().lastIndexOf("."))+"-"+targetLanguageIndex.getKey()+".xlsx";
                File trfile = trFilesMap.get(key);
                if(trfile!=null){
//                    System.out.println("目标语 index : "+targetLanguageIndex.getValue()+" 目标语言 mark : "+ExcelColumnsMark[targetLanguageIndex.getValue()] +" 目标语言 code:"+targetLanguageIndex.getKey());
                    //回稿文件存在则进行 数据回填
                    FileInputStream trFileInputStream = new FileInputStream(trfile.getPath());
                    XSSFWorkbook trWorkbook = new XSSFWorkbook(trFileInputStream);
                    trFileInputStream.close();
                    XSSFSheet trSheet = trWorkbook.getSheetAt(0);
                    if(trSheet.getLastRowNum()==sheet.getLastRowNum()){
                        //TODO 数据回填
                        for(int i=0;i<=sheet.getLastRowNum();i++){
                            //新翻文件 在
                            int trContentIndex=isNewTranslate?0:1;
                            Cell trCell=trSheet.getRow(i).getCell(trContentIndex);
                            Cell originalCell=sheet.getRow(i).getCell(targetLanguageIndex.getValue());

                            String trCellStr = getCellValueByCell(trCell);
                            String originalCellStr = getCellValueByCell(originalCell);
                            if (i==0){
                                //表头不为空并且相等，才进行后续的 数据回填
                                if(StringUtils.isNotEmpty(originalCellStr) && originalCellStr.equals(trCellStr)){
                                    continue;
                                }else{
                                    logger.error("原始文件表头标识 与 回稿文件 不一致请检查："+trfile.getName()+" 原始文件target index:"+targetLanguageIndex.getValue());
                                    logger.error("originalCellStr:"+originalCellStr+" trCellStr:" +trCellStr);
                                    trWorkbook.close();
                                    continue out1;
                                }
                            }
                            //以上条件都验证通过 直接做数据回填。
                            originalCell.setCellValue(trCellStr);

                            //TODO 20230919 新增对回写译文的检查， 检查转义符 如果译文包含  start
                            //获取目标语言 对应的原文
                            String originalSourceStr = "";
                            String originalSourceAppType = getCellValueByCell(sheet.getRow(i).getCell(sourceLanguageIndexMap.get("应用类型")));
                            // 如果应用类型不是服务端的 (目前非服务端的 是 ReTranslate文件里为空 和 NewTranslate 为 OS)都做检查
                            if(!"服务端".equalsIgnoreCase(originalSourceAppType)){
                                if("vzh_rHK".equalsIgnoreCase(targetLanguageIndex.getKey()) || "vzh_rTW".equalsIgnoreCase(targetLanguageIndex.getKey())){
                                    //查找原文（简中）所在列 index
                                    Cell originalSourceCell=sheet.getRow(i).getCell(sourceLanguageIndexMap.get("vzh_rCN"));
                                    originalSourceStr = getCellValueByCell(originalSourceCell);
                                }else{
                                    //查找原文（英文）所在列 index
                                    Cell originalSourceCell=sheet.getRow(i).getCell(sourceLanguageIndexMap.get("vus"));
                                    originalSourceStr = getCellValueByCell(originalSourceCell);
                                }
                                //检查译文非尖括号内的
                                String errMsg=checkConvertStr(originalSourceStr,trCellStr);
                                if(StringUtils.isNotEmpty(errMsg)){
                                    logger.error("\n转义符警告：译文文件:【"+trfile.getName()+"】\n" +
                                            "原文件：【"+file.getName()+"】\t列:\t"+ExcelColumnsMark[targetLanguageIndex.getValue()]+"\t行:\t"+(i+1)+"\t"+errMsg);
                                    logger.error("译文：【"+trCellStr+"】");
                                    isSuccess =false;
                                }
                            }
                            //TODO 20230919 新增对回写译文的检查， 检查转义符 如果译文包含  end
                        }
                    }else{
                        logger.error("回填数据过程，检测到文件行数不匹配,请检查文件:"+file.getName()+" 行数【"+sheet.getLastRowNum()+"】"+" 回稿文件："+trfile.getName()+" 行数【"+trSheet.getLastRowNum()+"】");
                        isSuccess=false;
                        //                        continue;
                    }
                    trWorkbook.close();
                }else{
                    logger.error("预期使用的回稿文件："+key +"不存在，请检查原文：["+file.getName()+"]的回稿文件是否完整!");
                    isSuccess=false;
                    continue;
                }
            }

            FileOutputStream excelFileOutPutStream = new FileOutputStream(file.getPath());
            // 将最新的 Excel 文件写入到文件输出流中，更新文件信息！
            workbook.write(excelFileOutPutStream);
            // 执行 flush 操作， 将缓存区内的信息更新到文件上
            excelFileOutPutStream.flush();
            // 使用后，及时关闭这个输出流对象， 好习惯，再强调一遍！
            excelFileOutPutStream.close();
            workbook.close();
        }


        //清空原文件及译文目录，方便下次直接 copy 文件不需要删除
        File deldir2 = new File (originalFielForder);
        File deldir3 = new File (translateFielForder);
        //TODO 清空输出目录
//        FileUtils.forceDelete(deldir2);
//        FileUtils.forceDelete(deldir3);


        logger.info("原始文件数量："+listOriginalFiles.size());
        logger.info("回稿文件数量："+listTranslateFiels.size());
        logger.info("合并后件数量："+newFileMap.size());

        processingRecords.put("批次数量",processingRecords.get("批次数量")+1);
        processingRecords.put("原始文件数量",processingRecords.get("原始文件数量")+listOriginalFiles.size());
        processingRecords.put("回稿文件数量",processingRecords.get("回稿文件数量")+listTranslateFiels.size());
        processingRecords.put("合并后件数量",processingRecords.get("合并后件数量")+newFileMap.size());

        for (File listOriginalFile : listOriginalFiles) {
            if(newFileMap.get(listOriginalFile.getName())==null){
                logger.error("原始文件： "+listOriginalFile.getName() +"未找到回稿文件!");
                isSuccess=false;
            }
        }
        return isSuccess;
    }


    //获取单元格各类型值，返回字符串类型
    private static String getCellValueByCell(Cell cell) {
        //判断是否为null或空串
        if (cell==null || cell.toString().trim().equals("")) {
            return "";
        }

        String cellValue = "";
        CellType cellType=cell.getCellType();
        if(cellType==CellType.FORMULA){ //表达式类型
            cellType=evaluator.evaluate(cell).getCellType();
        }else if(cellType==CellType.STRING){   //字符串类型
            cellValue= cell.getStringCellValue().trim();
            cellValue= StringUtils.isEmpty(cellValue) ? "" : cellValue;
        }else if(cellType==CellType.BOOLEAN){   //字符串类型
            cellValue = String.valueOf(cell.getBooleanCellValue());
        }else if(cellType==CellType.NUMERIC){   //字符串类型
            if (HSSFDateUtil.isCellDateFormatted(cell)) {  //判断日期类型
                cellValue =    DateFormatUtils.format(cell.getDateCellValue(), "yyyy-MM-dd");
            } else {  //否
                cellValue = new DecimalFormat("#.######").format(cell.getNumericCellValue());
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }

    public static  String checkConvertStr(String sourceStr,String targetStr){
//        target = "確定刪除所選取的<xliff:g id=\\\"plural\">%d</xliff:g>個標籤嗎？";
        // 增加空格为了 避免分割的内容在字符串最后 导致 比较时数量判断错误
        String target=" "+targetStr+" ";
        target=target.replaceAll("&lt;","<").replaceAll("&gt;",">");
        Pattern pattern = Pattern.compile("<([^<>]*)>");
        Matcher matcher = pattern.matcher(target);

        while (matcher.find()){
            String match = matcher.group();
            target=target.replace(match,"");
        }

        StringBuffer msg = new StringBuffer("");
        String [] shuangYinhaoNum=target.split("\"");
        String [] shuangYinhaoZhuanYiNum=target.split("\\\\\"");

        String [] danYinhaoNum=target.split("'");
        String [] danYinhaoZhuanYiNum=target.split("\\\\'");

        if(shuangYinhaoNum.length!=shuangYinhaoZhuanYiNum.length){
            msg.append("双引号缺失转义符 双引号数量:\t").append(shuangYinhaoNum.length-1).append("\t已转义数量:\t").append(shuangYinhaoZhuanYiNum.length-1);
        }
        if(danYinhaoNum.length!=danYinhaoZhuanYiNum.length){
            msg.append("\t单引号缺失转义符 单引号数量:\t").append(danYinhaoNum.length-1).append("\t已转义数量:\t").append(danYinhaoZhuanYiNum.length-1);
        }
//
        return msg.toString();
    }
}
