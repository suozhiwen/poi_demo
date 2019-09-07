package com.example.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.*;


public class Create {

    /*https://poi.apache.org/components/spreadsheet/quick-guide.html*/
    private static final String PATH = "D:/Excel/";

    public static void main(String[] args) {
        CreatingCells();
    }

    /**
     * 创建 两个不同后缀的excel
     *
     * @param
     */
    public static void NewWorkbook() {
        Workbook wb = new HSSFWorkbook();
        try (OutputStream fileOut = new FileOutputStream(new File(PATH + "workbook.xls"))) {
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        try (OutputStream fileOut = new FileOutputStream(new File(PATH + "workbook.xlsx"))) {
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    /**
     * 新表
     *
     * @param
     */
    public static void NewSheet() {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet sheet2 = wb.createSheet("second sheet");


        //请注意，工作表名称为Excel不得超过31个字符
        //并且不得包含以下任何字符：
        // 0x0000
        // 0x0003
        //冒号（:)
        //反斜杠（\）
        //星号（ *）
        //问号（？）
        //正斜杠（/）
        //打开方括号（[）
        //关闭方括号（]）

        //可以使用org.apache.poi.ss.util.WorkbookUtil #createSafeSheetName （String nameProposal）}
        //为了安全地创建有效名称，该实用程序用空格替换无效字符（''）
        String safeName = WorkbookUtil.createSafeSheetName("[0'Brien's sales*?]");  //返回“O'Brien's sales”
        Sheet sheet3 = wb.createSheet(safeName);

        try {
            OutputStream fileOut = new FileOutputStream(new File(PATH + "newSheet.xlsx"));
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    /**
     * 创建单元格
     */
    public static void CreatingCells(){
        Workbook wb = new HSSFWorkbook();
        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");

        //创建一行并在其中放入一些单元格 。 以0行为基础
        Row row = sheet.createRow(0);

        //创建一个单元格并在其中放置一个值。
        Cell cell = row.createCell(0);
        cell.setCellValue(1);

        //或者在一条线上做
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(creationHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);

        //将输出写入文件
        output.crate(wb,"CreatingCells");
    }
    /**
     * 创建日期单元格
     */
    public static void CreatingDateCells(){
        Workbook wb = new HSSFWorkbook();
        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");
        //创建一行并在其中放入一些单元格。以0行为基础
        Row row = sheet.createRow(0);
    }

}

class output{
    private static final String PATH = "D:/Excel/";
    public static void crate(Workbook wb,String path){
        try {
            OutputStream fileOut = new FileOutputStream(new File(PATH + path+".xls"));
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
