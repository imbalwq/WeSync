package com.xinyu.vivotool.vivo;


import com.xinyu.vivotool.ui.panel.StatusPanel;
import com.xinyu.vivotool.vivo.util.FileUtil;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
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
import java.util.stream.Collectors;

/**
 * @author liweiqing
 * @date 2023/9/1 16:05
 * @description vivo 新增多语言，并调整拆文件预处理规则
 */
public class VivoSplitExcel20231010 {

    private static FormulaEvaluator evaluator;
    private static Map<String,String> personLanguageCodeMap;
    private static final Logger logger = LoggerFactory.getLogger(VivoSplitExcel20231010.class);

    public static void main(String[] args) throws  Exception{
        //原文件目录 , 需要先手动解压文件, 压缩包暂时没有处理
        String folderPath="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo result\\原始文件";
        //处理后的译前文件
        String newFolderPath="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo result\\1、译前\\";
        //1、译前处理
        vivoExcelSplit(folderPath,newFolderPath);
    }

    /**
     *
     * @param originalFielForder   原始文件目录 客户原始文件
     * @param translateFielForder  翻译后的返回的excel文件目录
     * @param outputFolderPath     将合并后的数据输出到的目录
     */
    public static void excelDataCopyMerge(String originalFielForder, String translateFielForder,String outputFolderPath) throws IOException {

        originalFielForder="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo result\\2、译后回填处理\\Source";
        translateFielForder="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo result\\2、译后回填处理\\译文";
        outputFolderPath="C:\\Users\\Administrator\\Desktop\\temp\\zhaoxiaoling\\vivo result\\2、译后回填处理\\2、译后回填完成文件";

        //清空结果文件夹目录
        File deldir1 = new File (outputFolderPath);
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
                    }
                }
            }
        }

        //TODO 将回稿文件的数据 回填到 copy出来的原始文件中。
        //读取新文件检查需要回填的语言 有哪些，所有目标语言的回稿文件是否都存在
        for (Map.Entry<String, File> newfile : newFileMap.entrySet()) {
            File file=newfile.getValue();
            FileInputStream excelFileInputStreamTemp = new FileInputStream(file.getPath());
            XSSFWorkbook workbook = new XSSFWorkbook(excelFileInputStreamTemp);
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

            //记录该文件中 目标语言的索引位置
            Map<String,Integer> targetLanguageIndexMap = new HashMap<>();
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
                return;
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
            }
            //根据原文，检查回稿文件是否完整，  targetLanguageIndex:原文件 目标语言的索引位置
            out1:for (Map.Entry<String, Integer> targetLanguageIndex : targetLanguageIndexMap.entrySet()) {
                String key=file.getName().substring(0,file.getName().lastIndexOf("."))+"-"+targetLanguageIndex.getKey()+".xlsx";
                File trfile = trFilesMap.get(key);
                if(trfile!=null){
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
                                    System.out.println("originalCellStr:"+originalCellStr+" trCellStr:" +trCellStr);
                                    trWorkbook.close();
                                    continue out1;
                                }
                            }
                            //以上条件都验证通过 直接做数据回填。
                            originalCell.setCellValue(trCellStr);
                        }
                    }else{
                        logger.error("回填数据过程，检测到文件行数不匹配,请检查文件:"+file.getName()+" 回稿文件："+trfile.getName());
//                        continue;
                    }
                    trWorkbook.close();
                }else{
                    logger.error("预期使用的回稿文件："+key +"不存在，请检查原文：["+file.getName()+"]的回稿文件是否完整!");
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
        FileUtils.forceDelete(deldir2);
        FileUtils.forceDelete(deldir3);


        System.out.println("原始文件数量："+listOriginalFiles.size());
        System.out.println("回稿文件数量："+listTranslateFiels.size());
        System.out.println("合并后件数量："+newFileMap.size());

        for (File listOriginalFile : listOriginalFiles) {
            if(newFileMap.get(listOriginalFile.getName())==null){
                logger.error("原始文件： "+listOriginalFile.getName() +"未找到回稿文件!");
            }
        }
    }

    /**
     *
     * @param folderPath  原文件目录 , 需要先手动解压文件, 压缩包暂时没有处理
     * @param newFolderPath 处理后的译前文件存储位置
     * @throws Exception
     */
    public static String vivoExcelSplit(String folderPath,String newFolderPath) throws  Exception{
        //原文件目录 , 需要先手动解压文件, 压缩包暂时没有处理
        //清空结果文件目录
        File deldir1 = new File (newFolderPath);
        FileUtils.cleanDirectory(deldir1);
        String resultMsg ="";
        //初始化加载语言对与PM的对应关系
        personLanguageCodeMap();

        List<File> listFiles = FileUtil.readyAllFiel(folderPath,"xlsx");
        int i=0;
        StatusPanel.progressCurrent.setMaximum(listFiles.size());
        StatusPanel.progressCurrent.setValue(0);

        int errnum=0;
        for (File file : listFiles) {
            // Excel 数据拆分处理
            boolean suc=copySheet(file,newFolderPath);
            i++;
            StatusPanel.progressCurrent.setValue(i);
            //如果处理过程有问题则数量+1
            if(!suc){
                errnum+=1;
            }
        }
        if(errnum>0){
            return "拆分完成，但其中有"+errnum+"个文件拆分过程检测到问题，请查看【日志详情】！";
        }else{
            return "拆分文件完成";
        }
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


    /**
     * 根据 文件名是否包含 tReTranslate 及 newtranslate 划分一级目录 1、新翻  2、核对
     * 1、新翻  根据文件内应用模块一列属性值 划分为 OS、及服务端
     * 2、核对 根据要核对的语种 通过语言码 进行文件夹划分
     * @param fiel  译前source excel 文件
     * @param newFolderPath 处理后的文件存储路径
     * @throws IOException
     */
    public static boolean copySheet(File fiel, String newFolderPath) throws IOException {


        boolean isSuccess=true;
        //TODO 读取部分数据 用于判断 该文件应如何归类 及 目标语言表头的定位 start
        // 创建 Excel 文件的输入流对象
        FileInputStream excelFileInputStreamTemp = new FileInputStream(fiel.getPath());
        // XSSFWorkbook 就代表一个 Excel 文件
        // 创建其对象，就打开这个 Excel 文件
        XSSFWorkbook workbookTemp = null;
        ZipSecureFile.setMinInflateRatio(-1.0d);
        workbookTemp=new XSSFWorkbook(excelFileInputStreamTemp);
        // 输入流使用后，及时关闭！这是文件流操作中极好的一个习惯！
        excelFileInputStreamTemp.close();
        // XSSFSheet 代表 Excel 文件中的一张表格
        // 我们通过 getSheetAt(0) 指定表格索引来获取对应表格
        // 注意表格索引从 0 开始！
        XSSFSheet sheetTemp = workbookTemp.getSheetAt(0);
        //TODO 读取部分数据 用于判断 该文件应如何归类 及 目标语言表头的定位 end

        List<String> targetLanguageList = new ArrayList<>();
        //记录该文件中 目标语言的索引位置
        Map<String,Integer> targetLanguageIndexMap = new HashMap<>();

        //TODO 检查文件名 根据文件名，判断属于 新增翻译 还是 核对文件
        boolean isNewTranslate=true;
        String applicationType="";
        //检查原语言列 语言代码是否匹配
        int enSourceCodeColumnIndex=-1;
        //校正类excel 英文列 如果校正列有数据 使用校正列数据做为原文
        int enSourceCodeReviseColumnIndex=-1;
        int cnSourceCodeColumnIndex=-1;
        int sourceCodeColumnIndex=-1;
        //应用类型 在D列， 表头为："应用类型" 暂时固定
        int applicationTypeIndex=3;
        String newFolderPath2="";
        if(fiel.getName().toLowerCase().contains("retranslate")){
            //TODO execuetReTranslate()  业务暂时合并没做拆分
            isNewTranslate=false;
        }else if(fiel.getName().toLowerCase().contains("newtranslate")){
            //TODO execuetNewTranslate()    业务暂时合并没做拆分
            isNewTranslate=true;
        }else{
            logger.error("无法通过文件名确定是新增翻译或核对， 文件名为："+fiel.getName());
            logger.error("路径："+fiel.getPath());
            isSuccess=false;
            return false;
        }

        //TODO 后期修改为 从文件标题读取 判断原语言和目标语言 index start
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

        //2023年12月11日新增语言 start
        targetLanguageList.add("vmy_rMM");
        targetLanguageList.add("vmy_rZG");
        targetLanguageList.add("vms");
        targetLanguageList.add("vtr_rTR");
        //2023年12月11日新增语言 end

        //TODO 遍历表头确定语言码位置及原语言表头位置
        Row row0temp=sheetTemp.getRow(0);
        // 理论上表头不会存在重复的目标语言
        for(int columnIndex=0;columnIndex<row0temp.getLastCellNum();columnIndex++){
            String title=getCellValueByCell(row0temp.getCell(columnIndex)).trim();
            for (String s : targetLanguageList) {
                if(title.equalsIgnoreCase(s)){
                    targetLanguageIndexMap.put(title,columnIndex);
                }
            }
            if(title.equalsIgnoreCase("vus")){
                enSourceCodeColumnIndex=columnIndex;
                //TODO 20231124 客户模板修改后 核对文件 取消了 New vus一列， 当前临时使用 vus 列做为原文。
//                enSourceCodeReviseColumnIndex=columnIndex;
            }
            if("New vus".equalsIgnoreCase(title)){
                enSourceCodeReviseColumnIndex=columnIndex;
            }
            if(title.equalsIgnoreCase("vzh_rCN")){
                cnSourceCodeColumnIndex=columnIndex;
            }

        }
        if(enSourceCodeColumnIndex==-1){
            logger.error("无法通过表头 vus 定位到原语言所在列");
            workbookTemp.close();
            isSuccess=false;
            return false;
        }

        if(isNewTranslate){
            //判断新翻文件是 OS端 还是服务端
            applicationType=getCellValueByCell(sheetTemp.getRow(1).getCell(applicationTypeIndex));
            if(StringUtils.isEmpty(applicationType.trim())){
                logger.error("文件："+fiel.getPath() +"  [D列，应用类型为空]");
                isSuccess=false;
            }else{
                if(!applicationType.equalsIgnoreCase("OS") && !applicationType.equals("服务端")){
                    logger.error("文件："+fiel.getPath() +"  [应用类型无法确定:applicationType="+applicationType+"]");
                    isSuccess=false;
                }
            }
            newFolderPath2="新翻\\"+applicationType;
        }else{
            newFolderPath2="核对";
        }
        workbookTemp.close();

        //每个目标语言重新读取一次原文件做逻辑处理。
        for (Map.Entry<String, Integer> entry : targetLanguageIndexMap.entrySet()) {
            //记录当前文件是否有需要翻译的行，
            int needDataRowNum=0;
            // 创建 Excel 文件的输入流对象
            FileInputStream excelFileInputStream = new FileInputStream(fiel.getPath());
            // XSSFWorkbook 就代表一个 Excel 文件
            // 创建其对象，就打开这个 Excel 文件
            XSSFWorkbook workbookOriginal = new XSSFWorkbook(excelFileInputStream);
            // 输入流使用后，及时关闭！这是文件流操作中极好的一个习惯！
            excelFileInputStream.close();

            //TODO 因为客户原始的 word 不符合 openXml的规范 无法导入trados 双语格式，所以这里转存储一次  start
            XSSFWorkbook workbook = null;
            if(isNewTranslate){
                workbook=workbookOriginal;
            }else{
                //双语文件重新生成一个没有样式的 excel 原客户的 excel style 文件不符合微软 openXml 的规范，导入trados失败
//                workbook=workbookOriginal;
                //TODO 2023-11-29 发现copy后的 excel 填充的背景色有问题 暂时取消workbook的复制 start
                workbook=new XSSFWorkbook();
                copyExcelWorkBook(workbookOriginal,workbook);
                workbookOriginal.close();
                //TODO 2023-11-29 发现copy后的 excel 填充的背景色有问题 暂时取消workbook的复制 end
            }


            //TODO 因为客户原始的 word 不符合 openXml的规范 无法导入trados 双语格式，所以这里转存储一次  end

            XSSFSheet sheet = workbook.getSheetAt(0);
            // T列       U               v           w               x           y
            // 19       20              21          22             23          24
            // vnl  ,   vhr_rHR  ,   vfi_rFI  ,   vta_rIN  ,   vkk_rKZ  ,   ves_rUS
            XSSFRow rowO = sheet.getRow(0);
            int lastColunmNum=rowO.getLastCellNum();
            int rtColumnIndex=entry.getValue().intValue();       //ReTranslateColumnIndex
            if(rtColumnIndex>=lastColunmNum){
                workbook.close();
                continue;
            }
            XSSFCell langCodeTitleCell=sheet.getRow(0).getCell(rtColumnIndex);
            String langCode=getCellValueByCell(langCodeTitleCell);
            //TODO 2023年9月1日更新内容 如果 langCode=vzh_rHK 或 vzh_rTW 原文为 vzh_rCN
            if("vzh_rHK".equalsIgnoreCase(langCode) || "vzh_rTW".equalsIgnoreCase(langCode)){
                sourceCodeColumnIndex=cnSourceCodeColumnIndex;
            }else {
                if(isNewTranslate){
                    sourceCodeColumnIndex=enSourceCodeColumnIndex;
                }else{
                    sourceCodeColumnIndex=enSourceCodeReviseColumnIndex;
                }

            }
            if(targetLanguageIndexMap.containsKey(langCode)){
                //存储的路径为
                String fileName=fiel.getName().substring(0,fiel.getName().lastIndexOf("."));
                String newFileFolderPath="";
                if(personLanguageCodeMap.containsKey(langCode)){
                    //PM姓名， 目前暂时用 AB 代替
                    String personName=personLanguageCodeMap.get(langCode);
                    if("vzh_rHK".equalsIgnoreCase(langCode) || "vzh_rTW".equalsIgnoreCase(langCode)){
                        newFileFolderPath=newFolderPath+personName+"\\简中-其它\\"+newFolderPath2+"\\"+langCode+"\\"+fileName+"-"+langCode+".xlsx";
                    }else{
                        newFileFolderPath=newFolderPath+personName+"\\"+newFolderPath2+"\\"+langCode+"\\"+fileName+"-"+langCode+".xlsx";
                    }

                }else{
                    newFileFolderPath=newFolderPath+"未归属到任何pm待分配语言\\"+newFolderPath2+"\\"+langCode+"\\"+fileName+"-"+langCode+".xlsx";
                }

                //TODO 检查需要隐藏的行 和 列，核对文件只保存 S列(index 18) 英文(vus) 和 需要核对的语言列 (rtColumnIndex)
                //设置需要隐藏的列、行
//                rowO.setZeroHeight(true);
                String sourceCode=getCellValueByCell(sheet.getRow(0).getCell(sourceCodeColumnIndex));
                if(!sourceCode.toLowerCase().equals("vus") && !sourceCode.toLowerCase().equals("vzh_rcn") && !"New vus".equalsIgnoreCase(sourceCode)){
                    logger.error("文件："+fiel.getPath() +" 原语言列："+sourceCodeColumnIndex+"不匹配请检查,当前为："+sourceCode);
                    isSuccess=false;
                    workbook.close();
                    continue;
                }

                //设置需要隐藏的行 不包含标题行  rowIndex=1 表示从第二行开始
                for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                    // XSSFRow 代表一行数据
                    XSSFRow rowOriginal = sheet.getRow(rowIndex);
                    if (rowOriginal == null) {
                        logger.error("文件："+fiel.getPath() +" 行："+rowIndex+"为 null ，请检查!");
                        isSuccess=false;
                        continue;
                    }
                    //检查当前行是否需要隐藏
                    String rtColumnValue=getCellValueByCell(rowOriginal.getCell(rtColumnIndex));
                    String appTypeTemp=getCellValueByCell(rowOriginal.getCell(applicationTypeIndex));
                    if(isNewTranslate){
                        if(!appTypeTemp.equals(applicationType)){
                            logger.error("文件："+fiel.getPath() +" 行："+rowIndex+"应用类型为["+appTypeTemp+"],与其它行类型["+applicationType+"]不同，该文件目前需要手动处理！");
                            isSuccess=false;
                        }
                        //新翻文件并且有需要翻译的内容，第一次循环时 隐藏表头。
                        if(rowIndex==1 ){
                            sheet.getRow(0).setZeroHeight(true);
                        }
                        if(StringUtils.isNotEmpty(rtColumnValue)){
                            //如果不为空检查是否需要隐藏 隐藏标记有背景色的部分
                            short automatic=rowOriginal.getCell(rtColumnIndex).getCellStyle().getFillForegroundColor();
                            //需要翻译的部分 有内容，但是设置了背景色则 不翻译
                            if(automatic==IndexedColors.GREY_25_PERCENT.getIndex()){
                                rowOriginal.setZeroHeight(true);
                            }else{
                                //需要新翻，但是却有内容
                                logger.error("文件："+fiel.getPath() +"[行："+rowIndex+"列:"+rtColumnIndex+"] 需要新翻的目标语言不为【空】请检查");
                                isSuccess=false;
                                rowOriginal.setZeroHeight(true);
                            }
//                            if("-".equals(rtColumnValue.trim())){
//                                //不需要新翻
//                                rowOriginal.setZeroHeight(true);
//                            }else{
//                                //需要新翻，但是却有内容
//                                logger.error("文件："+fiel.getPath() +"[行："+rowIndex+"列:"+rtColumnIndex+"] 需要新翻的目标语言不为【空】请检查");
//                                rowOriginal.setZeroHeight(true);
//                            }
                        }else{
                            //需要翻译的部分虽然留空，但是设置了背景色则 不翻译
                            short automatic=rowOriginal.getCell(rtColumnIndex).getCellStyle().getFillForegroundColor();
                            //有背景色，无论什么颜色 都做隐藏处理
                            if(automatic==IndexedColors.GREY_25_PERCENT.getIndex()){
                                rowOriginal.setZeroHeight(true);
                            }
                        }
                    }else{
                        if(StringUtils.isEmpty(rtColumnValue)) {
                            short automatic=rowOriginal.getCell(rtColumnIndex).getCellStyle().getFillForegroundColor();
                            //有背景色，无论什么颜色 都做隐藏处理
                            if(automatic==IndexedColors.GREY_25_PERCENT.getIndex()){
                                rowOriginal.setZeroHeight(true);
                            }else{
                                logger.error("文件：" + fiel.getPath() + "[行：" + (rowIndex + 1) + "列:" + rtColumnIndex + "] 需要核对的内容为【空】请检查 "+"背景色："+automatic);
                                isSuccess=false;
                            }
//                        }else if("-".equals(rtColumnValue.trim())){
                        }else{
                            //如果不为空检查是否需要隐藏 隐藏标记有背景色的部分
                            short automatic=rowOriginal.getCell(rtColumnIndex).getCellStyle().getFillForegroundColor();
                            //有背景色，无论什么颜色 都做隐藏处理
                            if(automatic==IndexedColors.GREY_25_PERCENT.getIndex()){
                                rowOriginal.setZeroHeight(true);
                            }
                        }
                    }
                }

                //新翻插入1列，核对插入2列
                int addColumnNum=isNewTranslate?1:2;
                //从起始位置插入列
                sheet.shiftColumns(0,lastColunmNum,addColumnNum);
                //将原语言列复制到A列，目标核对语言列复制到B列
                int displayRowNum=0;
                for(int rowIndex=0;rowIndex<=sheet.getLastRowNum();rowIndex++){

                    Cell sourceStrCell=sheet.getRow(rowIndex).getCell(sourceCodeColumnIndex+addColumnNum);
                    Cell targetStrCell=sheet.getRow(rowIndex).getCell(rtColumnIndex+addColumnNum);
                    if(sheet.getRow(rowIndex).getZeroHeight()){
//                        System.out.println(fiel.getName()+"\t"+entry.getKey()+"\t"+sheet.getRow(rowIndex).getZeroHeight());
                    }else{
                        displayRowNum++;
                    }

                    if(isNewTranslate){
                        String  cellValue=getCellValueByCell(targetStrCell);
                        targetStrCell.getCellStyle().getFillBackgroundColor();
                        //如果是第一行就 复制目标语言 表头，如果是不需要翻译的 直接复制目标语言单元格
                        short automatic=targetStrCell.getCellStyle().getFillForegroundColor();
                        //有背景色，无论什么颜色 都做隐藏处理
                        if(rowIndex==0 || "-".equals(cellValue.trim()) || automatic==IndexedColors.GREY_25_PERCENT.getIndex()){
                            sheet.getRow(rowIndex).createCell(0).copyCellFrom(targetStrCell,new CellCopyPolicy());
                        }else{
                            //如果是需要翻译的 将原语言复制到该位置
                            sheet.getRow(rowIndex).createCell(0).copyCellFrom(sourceStrCell,new CellCopyPolicy());
                        }
                        //判断是否需要翻译，如果需要翻译设置
                    }else{
//                        if(enSourceCodeReviseColumnIndex!=-1){
//                            //查看英文校正列是否有数据，如果有数据 使用校正列的值做为原文
//                            Cell sourceReviseStrCell=sheet.getRow(rowIndex).getCell(enSourceCodeReviseColumnIndex+addColumnNum);
////                            //英文校正列 数据不为空时
////                            if(StringUtils.isNotEmpty(getCellValueByCell(sourceReviseStrCell))){
//                                //A列原语言
//                                sheet.getRow(rowIndex).createCell(0).copyCellFrom(sourceReviseStrCell,new CellCopyPolicy());
////                            }else {
////                                sheet.getRow(rowIndex).createCell(0).copyCellFrom(sourceStrCell,new CellCopyPolicy());
////                            }
//                        }else{
//                            //新翻文件 原文列
//                            sheet.getRow(rowIndex).createCell(0).copyCellFrom(sourceStrCell,new CellCopyPolicy());
//                        }
                        //A列原语言
                        sheet.getRow(rowIndex).createCell(0).copyCellFrom(sourceStrCell,new CellCopyPolicy());
                        //B列待核对的目标语言
                        sheet.getRow(rowIndex).createCell(1).copyCellFrom(targetStrCell,new CellCopyPolicy());
                    }
                }
                //
                //核对文件保留 A B列， 其它都设置为隐藏
                //新翻文件保留 A B列， 其它都设置为隐藏

                for(int c=addColumnNum;c<sheet.getRow(0).getLastCellNum();c++){
                    //设置列隐藏
                    sheet.setColumnHidden(c,true);
                }

//                    System.out.println("插入两行后最后一列index:"+sheet.getRow(0).getLastCellNum());
//                rtColumnIndex
//                sourceCodeColumnIndex
                //检查目录是否存在如果不存在先创建对应目录

                //如果 需要翻译的行数不为0则 创建新文件 因为发现客户有些文件 一条都不需要翻译。。。 但是又给了过来
                if(displayRowNum>0){
                File dir=new File(newFileFolderPath.substring(0,newFileFolderPath.lastIndexOf("\\")));
                dir.mkdirs();

                //进行存储

                    FileOutputStream excelFileOutPutStream = new FileOutputStream(newFileFolderPath);
                    // 将最新的 Excel 文件写入到文件输出流中，更新文件信息！
                    workbook.write(excelFileOutPutStream);
                    // 执行 flush 操作， 将缓存区内的信息更新到文件上
                    excelFileOutPutStream.flush();
                    // 使用后，及时关闭这个输出流对象， 好习惯，再强调一遍！
                    excelFileOutPutStream.close();
                }

                workbook.close();

            }else{
                logger.error("[范围外的语言代码，请检查列:"+rtColumnIndex +" 语言码："+langCode+"] 文件："+fiel.getPath());
                isSuccess=false;
            }
        }
        return isSuccess;
    }

    /**
     * 合并后的数据输出位置
     * @param translateFiel
     * @param originalFiel
     * @throws IOException
     */
    public static void excelDataMerge(File translateFiel, File originalFiel) throws IOException {
        //TODO
    }

    //TODO 将对应的翻译结果 复制到 对应的文件目录中
    public static boolean Move(String srcFile, String destPath)
    {
        // File (or directory) to be moved
        File file = new File(srcFile);
        File dir = new File(destPath);
        boolean success=false;
        //如果该路径下有文件则做 copy操作
        if(file.exists()){
            boolean createNewFoder=dir.mkdirs();
//            System.out.println("createNewFoder:"+createNewFoder);
            success= file.renameTo(new File(dir, file.getName()));
            if(!success){
                System.out.println("createNewFoder:"+createNewFoder+dir+file.getName());
            }
        }else{

        }
        return success;
    }


    /**
     * 注意 使用后 对 workbook 的关闭
     * @param workbookOriginal 原 workbook
     * @param newWorkbook 新 workbook
     * @throws IOException
     */
    public static void copyExcelWorkBook(XSSFWorkbook workbookOriginal,XSSFWorkbook newWorkbook) throws IOException {

        //TODO 因为客户原始的 word 不符合 openXml的规范 无法导入trados 双语格式，所以这里转存储一次  start
        XSSFSheet sheetOriginal = workbookOriginal.getSheetAt(0);
        evaluator=workbookOriginal.getCreationHelper().createFormulaEvaluator();
        int tempColumnNum=sheetOriginal.getRow(0).getLastCellNum();
        XSSFSheet newSheet = newWorkbook.createSheet(sheetOriginal.getSheetName());
        //遍历行
        for(int i=0;i<=sheetOriginal.getLastRowNum();i++){
            //遍历列
            XSSFRow rowi= newSheet.createRow(i);
            XSSFRow rowOriginal=sheetOriginal.getRow(i);
            for(int j=0;j<sheetOriginal.getRow(0).getLastCellNum();j++){

                //创建第一行第一个
                XSSFCell cell = rowi.createCell(j);
                CellCopyPolicy other = new CellCopyPolicy();
                //样式的 copy 经过实践发现 经常出问题，或者 影响trados 的导入。所以这里不copy style。
                other.setCopyCellStyle(false);
                cell.copyCellFrom(rowOriginal.getCell(j),other);
                CellStyle cellStyle=newWorkbook.createCellStyle();
                cellStyle.cloneStyleFrom(rowOriginal.getCell(j).getCellStyle());
                cell.setCellStyle(cellStyle);
            }
        }
//        workbookOriginal.close();
    }

    /**
     * 加载语言对与PM的分配情况 对应关系
     */
    public static void personLanguageCodeMap(){
        personLanguageCodeMap = new HashMap<>();
        //WY 简中
        personLanguageCodeMap.put("vzh_rHK","简中-其它");
        personLanguageCodeMap.put("vzh_rTW","简中-其它");

        //WY
        personLanguageCodeMap.put("ves_rUS","WY");
        personLanguageCodeMap.put("ves_rES","WY");
        personLanguageCodeMap.put("var_rEG","WY");

        //B PM
        personLanguageCodeMap.put("vtl","HY");
        personLanguageCodeMap.put("vpt_rPT","YY");
        personLanguageCodeMap.put("vfi_rFI","YY");
        personLanguageCodeMap.put("vkk_rKZ","YY");
        personLanguageCodeMap.put("vhr_rHR","YY");
        //SS
        personLanguageCodeMap.put("vta_rIN","HX");
        personLanguageCodeMap.put("vuk_rUA","HX");
        personLanguageCodeMap.put("vsr_rRS","HX");
        personLanguageCodeMap.put("vfa","HX");
        personLanguageCodeMap.put("vca_rES","HX");
        personLanguageCodeMap.put("vnl","HX");
        personLanguageCodeMap.put("vbn_rBD","HX");
        personLanguageCodeMap.put("vhi","HX");
        //ZY
        personLanguageCodeMap.put("vka","YY");
        personLanguageCodeMap.put("vmn_rMN","YY");
        personLanguageCodeMap.put("vpt_rBR","YY");
        personLanguageCodeMap.put("ven_rGB","HY");
        personLanguageCodeMap.put("vha","HY");

        //2023-12-11 新增四个语言
        personLanguageCodeMap.put("vmy_rMM","YT");
        personLanguageCodeMap.put("vmy_rZG","YT");
        personLanguageCodeMap.put("vms","YT");
        personLanguageCodeMap.put("vtr_rTR","HX");
    }
}
