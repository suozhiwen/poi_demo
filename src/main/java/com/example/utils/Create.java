package com.example.utils;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.*;
import java.util.Calendar;
import java.util.Date;


public class Create {

    /*https://poi.apache.org/components/spreadsheet/quick-guide.html*/
    private static final String PATH = "D:/Excel/";



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
    public static void CreatingCells() {
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
        output.crate(wb, "CreatingCells");
    }
    /**
     * 创建日期单元格
     */
    public static void CreatingDateCells() {
        Workbook wb = new HSSFWorkbook();
        CreationHelper creationHelper = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("new sheet");
        //创建一行并在其中放入一些单元格。以0行为基础
        Row row = sheet.createRow(0);

        //创建一个单元格并在其中添加日期值。第一个单元格没有样式
        //日期
        Cell cell = row.createCell(0);
        cell.setCellValue(new Date());


        //我们将第二个单元格设置为日期（和日间）
        //从工作薄中创建一个新的单元格样式
        //修改内置样式不仅影响这个单元格而且影响其他单元格
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("yyyy-MM-dd hh:mm:ss"));

        cell = row.createCell(1);
        System.out.println(new Date());
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);


        //你也可以将日期设置为 java.util.Calendar
        cell = row.createCell(2);
        System.out.println(Calendar.getInstance());
        cell.setCellValue(Calendar.getInstance());
        cell.setCellStyle(cellStyle);

        output.crate(wb, "CreatingDateCells");
    }
    /**
     * Working with different types of cells
     * 测试不同的数据类型
     */
    public static void WorkingWithDifferentTypesOfCells() {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet();
        Row row = sheet.createRow(2);
        row.createCell(0).setCellValue(1.1);
        row.createCell(1).setCellValue(new Date());
        row.createCell(2).setCellValue(Calendar.getInstance());
        row.createCell(3).setCellValue("a string");
        row.createCell(4).setCellValue(true);
        System.out.println(CellType.ERROR);
        row.createCell(5).setCellType(CellType._NONE);
        output.crate(wb, "WorkingWithDifferentTypesOfCells");
    }
    /**
     * 演示各种对齐选项
     * Demonstrates various alignment options
     */
    public static void DemonstratesVariousAlignmentOptions() {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet();
        //从 2+1 行开始
        Row row = sheet.createRow(2);
        row.setHeightInPoints(30);


        /**
         * HorizontalAlignment
         * http://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/HorizontalAlignment.html
         *
         * VerticalAlignment
         *http://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/VerticalAlignment.html
         */
        createCell(wb, row, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM);
        createCell(wb, row, 1,HorizontalAlignment.CENTER_SELECTION,VerticalAlignment.CENTER);
        createCell(wb, row, 2,HorizontalAlignment.FILL,VerticalAlignment.CENTER);
        createCell(wb, row, 3,HorizontalAlignment.FILL,VerticalAlignment.CENTER);
        createCell(wb, row, 4,HorizontalAlignment.JUSTIFY,VerticalAlignment.JUSTIFY);
        createCell(wb, row, 5,HorizontalAlignment.LEFT,VerticalAlignment.TOP);
        createCell(wb, row, 6,HorizontalAlignment.RIGHT,VerticalAlignment.TOP);

        output.crate(wb,"DemonstratesVariousAlignmentOptions");
    }
    /**
     * 创建一个单元格并以某种方式对其它
     *
     * @param wb     工作簿
     * @param row    行
     * @param column 列
     * @param halign 创建单元格 halign 单元格水平对齐
     * @param valign 单元格的垂直对齐方式
     */
    private static void createCell(Workbook wb, Row row, int column, HorizontalAlignment halign, VerticalAlignment valign) {
        Cell cell = row.createCell(column);
        cell.setCellValue("Align It");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(halign);
        cellStyle.setVerticalAlignment(valign);
        cell.setCellStyle(cellStyle);
    }

    /**
     * 迭代行和单元格
     * Fills and colors
     */
    public static void FillsAndColors(){



        output.crate(wb,"IterateOverRowsAndCells");
    }



    public static void main(String[] args) {
        /*NewWorkbook();*/
        /*NewSheet();*/
        /*CreatingCells();*/
        /*CreatingDateCells();*/
        /*WorkingWithDifferentTypesOfCells();*/
        /*DemonstratesVariousAlignmentOptions();*/
        IterateOverRowsAndCells();
    }
}

class output {
    private static final String PATH = "D:/Excel/";

    public static void crate(Workbook wb, String path) {
        try {
            OutputStream fileOut = new FileOutputStream(new File(PATH + path + ".xls"));
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
