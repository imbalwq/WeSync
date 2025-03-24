package com.luoboduner.wesync.tools;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.lang.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;

/**
 * @author liweiqing
 * @date 2024/1/8 12:57
 * @description
 */
public class ExcelRead {


    private static FormulaEvaluator evaluator;
    /**
     * 系统当前路径
     */
    private final static String CURRENT_DIR = System.getProperty("user.dir");

    public static void main(String[] args) throws IOException {
        String origianl= CURRENT_DIR+File.separator+"src"+File.separator+"main"+File.separator+"resources"+File.separator+"zhangxinwen"+File.separator + "线下语言码对照关系.xlsx";
        readExcelData(origianl);
    }

    //
    public static String[][] readExcelData(String absPath) throws IOException {
//        String origianl= CURRENT_DIR+File.separator+"src"+File.separator+"main"+File.separator+"resources"+File.separator+"zhangxinwen"+File.separator + "线下语言码对照关系.xlsx";
        File origianlFile = new File (absPath);

        FileInputStream excelFileInputStream = new FileInputStream(origianlFile.getPath());
        // XSSFWorkbook 就代表一个 Excel 文件
        // 创建其对象，就打开这个 Excel 文件
        XSSFWorkbook workbook = new XSSFWorkbook(excelFileInputStream);
        // 输入流使用后，及时关闭！这是文件流操作中极好的一个习惯！
        excelFileInputStream.close();
        //遍历原文件的所有 工作簿
        XSSFSheet sheet=workbook.getSheetAt(0);

        evaluator=workbook.getCreationHelper().createFormulaEvaluator();
        String[][] data = new String[sheet.getLastRowNum()+1][100];
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if(row==null){
                System.err.println("r:"+r);
            }
            for(int l=0;l<row.getLastCellNum();l++){
                Cell cell=row.getCell(l);
                String str = getCellValueByCell(cell);
                data[r][l]=str;
            }
        }
//        for (int i=0;i<data.length;i++){
//            for (int j=0;j<data[i].length;j++) {
//                System.out.print(data[i][j]+"\t");
//            }
//            System.out.println("");
//        }
        workbook.close();
        return data;
    }

    public static String[][] readExcelData(String absPath,String sheetName) throws IOException {
//        String origianl= CURRENT_DIR+File.separator+"src"+File.separator+"main"+File.separator+"resources"+File.separator+"zhangxinwen"+File.separator + "线下语言码对照关系.xlsx";
        File origianlFile = new File (absPath);

        FileInputStream excelFileInputStream = new FileInputStream(origianlFile.getPath());
        // XSSFWorkbook 就代表一个 Excel 文件
        // 创建其对象，就打开这个 Excel 文件
        XSSFWorkbook workbook = new XSSFWorkbook(excelFileInputStream);
        // 输入流使用后，及时关闭！这是文件流操作中极好的一个习惯！
        excelFileInputStream.close();
        //遍历原文件的所有 工作簿
        XSSFSheet sheet=null;
        if (StringUtils.isNotEmpty(sheetName)) {
           sheet=workbook.getSheet(sheetName);
        }else{
            sheet=workbook.getSheetAt(0);
        }

        evaluator=workbook.getCreationHelper().createFormulaEvaluator();
        String[][] data = new String[sheet.getLastRowNum()+1][255];
        out:for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if(row==null){
//                System.err.println("r:"+r);
                continue out;
            }
            for(int l=0;l<row.getLastCellNum();l++){
                Cell cell=row.getCell(l);
                String str = getCellValueByCell(cell);
                data[r][l]=str;
            }
        }
//        for (int i=0;i<data.length;i++){
//            for (int j=0;j<data[i].length;j++) {
//                System.out.print(data[i][j]+"\t");
//            }
//            System.out.println("");
//        }
        workbook.close();
        return data;
    }

    //获取单元格各类型值，返回字符串类型
    public static String getCellValueByCell(Cell cell) {
        //判断是否为null或空串
        if (cell==null || cell.toString().trim().equals("")) {
            return "";
        }
        String cellValue = "";
        CellType cellType=cell.getCellType();
        if(cellType==CellType.FORMULA){ //表达式类型
            CellType cvalueType=evaluator.evaluate(cell).getCellType();
            if(cvalueType==CellType.STRING){   //字符串类型
                cellValue= cell.getStringCellValue().trim();
                cellValue= StringUtils.isEmpty(cellValue) ? "" : cellValue;
            }else if(cvalueType==CellType.BOOLEAN){   //字符串类型
                cellValue = String.valueOf(cell.getBooleanCellValue());
            }else if(cvalueType==CellType.NUMERIC){   //字符串类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {  //判断日期类型
                    cellValue =    DateFormatUtils.format(cell.getDateCellValue(), "yyyy-MM-dd");
                } else {  //否
                    cellValue = new DecimalFormat("#.######").format(cell.getNumericCellValue());
                }
            }else{
                cellValue = "";
            }

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
}
