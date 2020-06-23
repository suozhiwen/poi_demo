package com.example.utils;

import com.example.entity.User;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.*;
import java.util.*;


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
        createCell(wb, row, 1, HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.CENTER);
        createCell(wb, row, 2, HorizontalAlignment.FILL, VerticalAlignment.CENTER);
        createCell(wb, row, 3, HorizontalAlignment.FILL, VerticalAlignment.CENTER);
        createCell(wb, row, 4, HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY);
        createCell(wb, row, 5, HorizontalAlignment.LEFT, VerticalAlignment.TOP);
        createCell(wb, row, 6, HorizontalAlignment.RIGHT, VerticalAlignment.TOP);

        output.crate(wb, "DemonstratesVariousAlignmentOptions");
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
     * 填充和颜色
     * Fills and colors
     */
    public static void FillsAndColors() {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");
        //创建一行并在其中放入一些单元格 。以0行为基础
        Row row = sheet.createRow(1);
        //设置样式
        CellStyle style = wb.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);
        Cell cell = row.createCell(1);
        cell.setCellValue("x");
        cell.setCellStyle(style);

        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell = row.createCell(2);
        cell.setCellValue("X");
        cell.setCellStyle(style);

        output.crate(wb, "FillsAndColors");
    }

    /**
     * 合并细胞
     * Merging cells
     */
    public static void MergingCells() {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet();
        Row row = sheet.createRow(1);
        sheet.addMergedRegion(new CellRangeAddress(
                3, //first row (0-based)           竖列开始的格子
                4, //last row  (0-based)           竖列结束的格子
                0, //first column (0-based)         横列开始的格子
                0  //last column  (0-based)         横列结束的格子
        ));

        sheet.addMergedRegion(new CellRangeAddress(
                1, //first row (0-based)           竖列开始的格子
                3, //last row  (0-based)           竖列结束的格子
                1, //first column (0-based)         横列开始的格子
                4  //last column  (0-based)         横列结束的格子
        ));

        output.crate(wb, "MergingCells");
    }


    public static void test() {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet();
        //行
        Row titleRow = null;
        //单元格
        Cell titleCell;
        sheet.addMergedRegion(new CellRangeAddress(0, 2, 0, 4));
        titleRow = sheet.createRow(0);
        titleCell = titleRow.createCell(0);
        /* titleCell.setCellType(CellType.STRING);*/
        titleCell.setCellValue("2018年度能源科技进步奖");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        titleCell.setCellStyle(cellStyle);

        sheet.addMergedRegion(new CellRangeAddress(3, 6, 0, 1));
        titleRow = sheet.createRow(3);
        titleCell = titleRow.createCell(0);
        titleCell.setCellValue("测试竖列");
        CellStyle cellStyle2 = wb.createCellStyle();
        //文字旋转
        cellStyle2.setRotation((short) 255);
        titleCell.setCellStyle(cellStyle2);


        sheet.addMergedRegion(new CellRangeAddress(3, 4, 2, 5));
        titleCell = titleRow.createCell(2);
        sheet.createRow(4);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("测试竖列1");
        CellStyle cellStyle1 = wb.createCellStyle();
        cellStyle1.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyle1);


        sheet.addMergedRegion(new CellRangeAddress(3, 4, 6, 10));
        titleCell = titleRow.createCell(6);
        sheet.createRow(4);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("测试竖列2");
        CellStyle cellStyle5 = wb.createCellStyle();
        cellStyle5.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyle5.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyle5);


        String[] title = {"用户名称", "年龄", "密码"};

        //创建数据
        String[][] content = CreateData(title);

        //行
        titleRow = sheet.createRow(7);


        /**
         * 从第 i+1 的竖列开始
         * titleRow.createCell(i+1)
         *
         */
        //创建标题
        for (int i = 0; i < title.length; i++) {
            titleCell = titleRow.createCell(i + 1);
            titleCell.setCellValue(title[i]);
        }

        //创建内容
        for (int i = 0; i < content.length; i++) {
            titleRow = sheet.createRow(i + 8);
            for (int j = 0; j < content[i].length; j++) {
                //将内容按顺序赋给对应的列对象
                titleRow.createCell(j + 1).setCellValue(content[i][j]);
            }
        }

        output.crate(wb, "test");
    }

    //创建数据
    public static String[][] CreateData(String[] title) {
        List<Map<String, Object>> lists = new ArrayList<>();

        Map<String, Object> map = new HashMap<>();
        map.put("id", 1);
        map.put("name", "名称1");
        map.put("model", "规格型号1");
        map.put("number", 100);
        map.put("unitPrice", "500");
        map.put("brand", "华为");
        map.put("unit", "台");
        map.put("supplier", "徐州华为总经销");
        map.put("remark", "自己人");
        map.put("subtotal", 10000);
        lists.add(map);

        Map<String, Object> map1 = new HashMap<>();
        map1.put("id", 1);
        map1.put("name", "名称2");
        map1.put("model", "规格型号2");
        map1.put("number", 100);
        map1.put("unitPrice", "500");
        map1.put("brand", "华为");
        map1.put("unit", "台");
        map1.put("supplier", "徐州华为总经销");
        map1.put("remark", "自己人");
        map1.put("subtotal", 10000);
        lists.add(map1);

        Map<String, Object> map2 = new HashMap<>();
        map2.put("id", 1);
        map2.put("name", "名称3");
        map2.put("model", "规格型号4");
        map2.put("number", 100);
        map2.put("unitPrice", "500");
        map2.put("brand", "华为");
        map2.put("unit", "台");
        map2.put("supplier", "徐州华为总经销");
        map2.put("remark", "自己人");
        map2.put("subtotal", 10000);
        lists.add(map2);

        String[][] content = new String[lists.size()][title.length];
        for (int i = 0; i < lists.size(); i++) {
            content[i] = new String[title.length];
            Map<String, Object> object = lists.get(i);
            content[i][0] = object.get("id").toString();
            content[i][1] = object.get("name").toString();
            content[i][2] = object.get("model").toString();
            content[i][3] = object.get("number").toString();
            content[i][4] = object.get("unitPrice").toString();
            content[i][5] = object.get("brand").toString();
            content[i][6] = object.get("unit").toString();
            content[i][7] = object.get("supplier").toString();
            content[i][8] = object.get("remark").toString();
            content[i][9] = object.get("subtotal").toString();
        }


        return content;
    }

    public static void excel() {

        List<Map<String, Object>> lists = new ArrayList<>();

        String[] title = {"用户名称", "年龄", "密码"};
        Map<String, Object> map = new HashMap<>();
        map.put("name", "李四");
        map.put("age", 12);
        lists.add(map);

        Map<String, Object> map1 = new HashMap<>();
        map1.put("name", "李四1");
        map1.put("age", 124);
        lists.add(map1);

        Map<String, Object> map2 = new HashMap<>();
        map2.put("name", "李四2");
        map2.put("age", 12000);
        lists.add(map2);

        String[][] content = new String[lists.size()][title.length];


        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet();
        //行
        Row titleRow = sheet.createRow(0);
        //单元格
        Cell titleCell = null;

        //创建标题
        for (int i = 0; i < title.length; i++) {
            titleCell = titleRow.createCell(i);
            titleCell.setCellValue(title[i]);
        }

        for (int i = 0; i < lists.size(); i++) {
            content[i] = new String[title.length];
            Map<String, Object> object = lists.get(i);
            if (object.containsKey("name")) {
                content[i][0] = object.get("name").toString();
            }
            if (object.containsKey("age")) {
                content[i][1] = object.get("age").toString();
            }
        }

        //创建内容
        for (int i = 0; i < content.length; i++) {
            titleRow = sheet.createRow(i + 1);
            for (int j = 0; j < content[i].length; j++) {
                //将内容按顺序赋给对应的列对象
                titleRow.createCell(j).setCellValue(content[i][j]);
            }
        }

        output.crate(wb, "test1");
    }

    /**
     * 采购计划
     */
    public static void purchasePlan() {
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet();
        //行
        Row titleRow = null;
        //单元格
        Cell titleCell;
        //sheet.addMergedRegion(new CellRangeAddress(0, 2, 0, 4));
//        titleRow = sheet.createRow(0);
//        titleCell = titleRow.createCell(0);
//        /* titleCell.setCellType(CellType.STRING);*/
//        titleCell.setCellValue("2018年度能源科技进步奖");
//        CellStyle cellStyle = wb.createCellStyle();
//        cellStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION);
//        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
//
//        titleCell.setCellStyle(cellStyle);
//
//        sheet.addMergedRegion(new CellRangeAddress(3, 6, 0, 1));
//        titleRow = sheet.createRow(3);
//        titleCell = titleRow.createCell(0);
//        titleCell.setCellValue("测试竖列");
//        CellStyle cellStyle2 = wb.createCellStyle();
//        //文字旋转
//        cellStyle2.setRotation((short)255);
//        titleCell.setCellStyle(cellStyle2);

        /**
         * addMergedRegion 合并单元格
         * @param firstRow   第一行索引
         * @param lastRow    最后一行(包括在内)的索引，必须等于或大于(第一行索引)
         * @param firstCol   第一列索引
         * @param lastCol   最后一列(包括在内)的索引，必须等于或大于 第一列索引
         */
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 1));
        titleRow = sheet.createRow(0);
        titleCell = titleRow.createCell(0);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("计划单号");
        CellStyle cellStyle1 = wb.createCellStyle();
        cellStyle1.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyle1);


        sheet.addMergedRegion(new CellRangeAddress(0, 1, 2, 7));
        titleCell = titleRow.createCell(2);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("CGJH2020000326");
        CellStyle cellStyle5 = wb.createCellStyle();
        //左对齐
        cellStyle5.setAlignment(HorizontalAlignment.LEFT);
        //垂直居中
        cellStyle5.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyle5);


        //每当切换行就要新生成一个在同一行构建单元格则不需要
        Row titleRow1 = null;
        //单元格
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 0, 1));
        //创建新的行对象
        titleRow1 = sheet.createRow(2);
        titleCell = titleRow1.createCell(0);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("申请人");
        CellStyle cellStyle3 = wb.createCellStyle();
        cellStyle3.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyle3.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyle3);

        //单元格
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 2, 3));
        //创建新的行对象
        titleCell = titleRow1.createCell(2);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("延续");
        CellStyle cellStyle4 = wb.createCellStyle();
        cellStyle4.setAlignment(HorizontalAlignment.LEFT);
        cellStyle4.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyle4);

        //申请时间标题
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 4, 5));
        //创建新的行对象
        titleCell = titleRow1.createCell(4);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("申请时间");
        CellStyle cellStyleDateTitle = wb.createCellStyle();
        cellStyleDateTitle.setAlignment(HorizontalAlignment.LEFT);
        cellStyleDateTitle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleDateTitle);

        //申请时间值
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 6, 7));
        //创建新的行对象
        titleCell = titleRow1.createCell(6);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("2020-06-23 14:51:32");
        CellStyle cellStyleDateValue = wb.createCellStyle();
        cellStyleDateValue.setAlignment(HorizontalAlignment.LEFT);
        cellStyleDateValue.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleDateValue);


        //标题
        Row titleRow2 = null;
        //单元格
        sheet.addMergedRegion(new CellRangeAddress(4, 5, 0, 1));
        //创建新的行对象
        titleRow2 = sheet.createRow(4);
        titleCell = titleRow2.createCell(0);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("标题");
        CellStyle cellStyleApplyTitle = wb.createCellStyle();
        cellStyleApplyTitle.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyleApplyTitle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleApplyTitle);

        //标题值
        //单元格
        sheet.addMergedRegion(new CellRangeAddress(4, 5, 2, 7));
        //创建新的行对象
        titleCell = titleRow2.createCell(2);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("dhso");
        CellStyle cellStyleApplyTitleValue = wb.createCellStyle();
        cellStyleApplyTitleValue.setAlignment(HorizontalAlignment.LEFT);
        cellStyleApplyTitleValue.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleApplyTitleValue);


        //第四行行采购类型和采购周期
        Row titleRow3 = null;
        //单元格 采购类型
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 0, 1));
        //创建新的行对象
        //从第几行开始创建
        titleRow3 = sheet.createRow(6);
        //从第几列开始创建
        titleCell = titleRow3.createCell(0);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("采购类型");
        CellStyle cellStyleApplyType = wb.createCellStyle();
        cellStyleApplyType.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyleApplyType.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleApplyType);

        //设备采购
        //单元格
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 2, 3));
        //创建新的行对象
        titleCell = titleRow3.createCell(2);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("设备采购");
        CellStyle cellStyleApplyTypeValue = wb.createCellStyle();
        cellStyleApplyTypeValue.setAlignment(HorizontalAlignment.LEFT);
        cellStyleApplyTypeValue.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleApplyTypeValue);

        //采购周期
        //单元格
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 4, 5));
        //创建新的行对象
        titleCell = titleRow3.createCell(4);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("采购周期");
        CellStyle cellStyleApplyPeriod = wb.createCellStyle();
        cellStyleApplyPeriod.setAlignment(HorizontalAlignment.LEFT);
        cellStyleApplyPeriod.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleApplyPeriod);


        //单元格1年第3季度
        sheet.addMergedRegion(new CellRangeAddress(6, 7, 6, 7));
        //创建新的行对象
        titleCell = titleRow3.createCell(6);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("1年第3季度");
        CellStyle cellStyleApplyPeriodValue = wb.createCellStyle();
        cellStyleApplyPeriodValue.setAlignment(HorizontalAlignment.LEFT);
        cellStyleApplyPeriodValue.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleApplyPeriodValue);


        //第四行行采购类型和采购周期
        Row titleRow5 = null;
        //单元格 采购类型
        sheet.addMergedRegion(new CellRangeAddress(8, 9, 0, 1));
        //创建新的行对象
        //从第几行开始创建
        titleRow5 = sheet.createRow(8);
        //从第几列开始创建
        titleCell = titleRow5.createCell(0);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("备注");
        CellStyle cellStyleRemark = wb.createCellStyle();
        cellStyleRemark.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyleRemark.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleRemark);

        //单元格1年第3季度
        sheet.addMergedRegion(new CellRangeAddress(8, 9, 2, 7));
        //创建新的行对象
        titleCell = titleRow5.createCell(2);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("122312323123123123");
        CellStyle cellStyleRemarkValue = wb.createCellStyle();
        cellStyleRemarkValue.setAlignment(HorizontalAlignment.LEFT);
        cellStyleRemarkValue.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleRemarkValue);


        String[] title = {"序号", "名称", "规格型号", "数量", "预计单价", "品牌", "单位", "供应商", "备注", "小计"};

        //创建数据
        String[][] content = CreateData(title);

        //行
        titleRow = sheet.createRow(10);


        /**
         * 从第 i+1 的竖列开始
         * titleRow.createCell(i+1)
         *
         */
        //创建标题
        for (int i = 0; i < title.length; i++) {
            titleCell = titleRow.createCell(i);
            titleCell.setCellValue(title[i]);
        }

        //创建内容
        for (int i = 0; i < content.length; i++) {
            titleRow = sheet.createRow(i + 11);
            for (int j = 0; j < content[i].length; j++) {
                //将内容按顺序赋给对应的列对象
                titleRow.createCell(j).setCellValue(content[i][j]);
            }
        }

        int row = content.length+1+10;
        //第四行行采购类型和采购周期
        Row titleRow6 = null;
        //单元格 采购类型
        sheet.addMergedRegion(new CellRangeAddress(row, row+1, 0, 0));
        //创建新的行对象
        //从第几行开始创建
        titleRow6 = sheet.createRow(row);
        //从第几列开始创建
        titleCell = titleRow6.createCell(0);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("合计");
        CellStyle cellStyleTotal = wb.createCellStyle();
        cellStyleTotal.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyleTotal.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleTotal);


        //单元格 采购类型
        sheet.addMergedRegion(new CellRangeAddress(row, row+1, 1, 3));
        //创建新的行对象
        //从第几列开始创建
        titleCell = titleRow6.createCell(1);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("1533");
        CellStyle cellStyleTotal1 = wb.createCellStyle();
        cellStyleTotal1.setAlignment(HorizontalAlignment.RIGHT);
        cellStyleTotal1.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleTotal1);


        //单元格 采购类型
        sheet.addMergedRegion(new CellRangeAddress(row, row+1, 4, 9));
        //创建新的行对象
        //从第几列开始创建
        titleCell = titleRow6.createCell(4);
        titleCell.setCellType(CellType.STRING);
        titleCell.setCellValue("117585");
        CellStyle cellStyleTotal2 = wb.createCellStyle();
        cellStyleTotal2.setAlignment(HorizontalAlignment.RIGHT);
        cellStyleTotal2.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCell.setCellStyle(cellStyleTotal2);

        output.crate(wb, "purchasePlan");
    }

    public static void main(String[] args) {
        /*NewWorkbook();*/
        /*NewSheet();*/
        /*CreatingCells();*/
        /*CreatingDateCells();*/
        /*WorkingWithDifferentTypesOfCells();*/
        /*DemonstratesVariousAlignmentOptions();*/
        /*FillsAndColors();*/
        /* MergingCells();*/
//        test();
        purchasePlan();
        /*excel();*/
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
