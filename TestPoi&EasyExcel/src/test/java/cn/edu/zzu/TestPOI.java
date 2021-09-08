/*
package cn.edu.zzu;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.*;
import java.util.Date;

public class TestPOI {

    */
/**
     * HSSFWorkbook excel文档对象
     * <p>
     * HSSFSheet excel的sheet HSSFRow excel的行
     * <p>
     * HSSFCell excel的单元格 HSSFFont excel字体
     * <p>
     * HSSFName 名称 HSSFDataFormat 日期格式
     * <p>
     * HSSFHeader sheet头
     * <p>
     * HSSFFooter sheet尾
     * <p>
     * HSSFCellStyle cell样式
     * <p>
     * HSSFDateUtil 日期
     * <p>
     * HSSFPrintSetup 打印
     * <p>
     * HSSFErrorConstants 错误信息表
     *//*


    String PATH = "D:\\Java\\idea_workspace\\TestPoi&EasyExcel\\";



    */
/*======================读=============================*//*


    */
/**
     * 03 --  读取一行一列
     *
     * @throws Exception
     *//*

    @Test
    public void testRead03() throws Exception {
        FileInputStream fileInputStream = new FileInputStream(PATH + "粉丝统计.xls");

        HSSFWorkbook sheets = new HSSFWorkbook(fileInputStream);

        HSSFSheet sheetAt = sheets.getSheetAt(0);

        //读取一行一列
        HSSFRow row = sheetAt.getRow(0);
        HSSFCell cell = row.getCell(0);

        System.out.println(cell.getStringCellValue());

        fileInputStream.close();
    }

    */
/**
     * 07 -- 读取一行一列
     *//*


    @Test
    public void testRead07() throws Exception {
        FileInputStream fileInputStream = new FileInputStream(PATH + "粉丝统计.xlsx");

        XSSFWorkbook sheets = new XSSFWorkbook(fileInputStream);

        XSSFSheet sheetAt = sheets.getSheetAt(0);

        XSSFRow row = sheetAt.getRow(0);
        XSSFCell cell = row.getCell(0);

        System.out.println(cell.getStringCellValue());

        fileInputStream.close();
    }

    @Test
    public void testCellType() throws Exception {
        FileInputStream fileInputStream = new FileInputStream(PATH + "会员消费商品明细表.xls");

        HSSFWorkbook sheets = new HSSFWorkbook(fileInputStream);

        HSSFSheet sheetAt = sheets.getSheetAt(0);

        //读取标题所有内容
        HSSFRow row = sheetAt.getRow(0);
        if (row != null) {//行不空

            int cellNum = row.getPhysicalNumberOfCells();
            for (int i = 0; i < cellNum; i++) {
                HSSFCell cell = row.getCell(i);
                if (cell != null) {
                    CellType cellType = cell.getCellType();
                    String stringCellValue = cell.getStringCellValue();
                    System.out.print(stringCellValue +" | ");
                }
            }
            System.out.println();
        }

        //读取其他行
        //行数
        int rowNum = sheetAt.getPhysicalNumberOfRows();
        for (int i = 1; i < rowNum; i++) {
            HSSFRow row1 = sheetAt.getRow(i);
            if (row1 != null) {

                //单元格数
                int cellNum = row1.getPhysicalNumberOfCells();

                for (int j = 0; j < cellNum; j++) {

                    System.out.print("[ " + (i + 1) + "-" + (j + 1) + " ]");
                    HSSFCell cell = row1.getCell(j);

                    if (cell != null) {
                        CellType cellType = cell.getCellType();
                        String cellValue = "";
                        switch (cellType) {
                            case STRING:
                                cellValue = cell.getStringCellValue();
                                System.out.print(" [STRING] ");
                                break;
                            case BOOLEAN:
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                System.out.print(" [BOOLEAN] ");
                                break;
                            case BLANK://空
                                cellValue = cell.getStringCellValue();
                                System.out.print(" [BLANK] ");
                                break;
                            case NUMERIC:
                                boolean cellDateFormatted = DateUtil.isCellDateFormatted(cell);
                                if(cellDateFormatted){
                                    System.out.println("[日期]");
                                    Date dateCellValue = cell.getDateCellValue();
                                    cellValue = new DateTime(dateCellValue).toString("yyyy-MM-dd HH:mm:ss");
                                }else{
                                    // 不是日期格式，则防止当数字过长时以科学计数法显示
                                    System.out.println("[转成字符串]");
                                    cell.setCellType(CellType.STRING);
                                    cellValue = cell.toString();
                                }
                                System.out.print(" [NUMERIC] ");
                                break;
                            case ERROR:
                                System.out.print(" [ERROR] : 类型错误！");
                                break;
                        }
                        System.out.println("  "+cellValue);
                    }

                }
            }
            System.out.println("=============================");
        }

    }


    */
/**
     * 读取计算公式
     *
     * @throws IOException
     *//*

    @Test
    public void testFormula() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "计算公式.xls");

        HSSFWorkbook sheets = new HSSFWorkbook(fileInputStream);

        HSSFSheet sheetAt = sheets.getSheetAt(0);

        HSSFRow row = sheetAt.getRow(4);

        HSSFCell cell = row.getCell(0);

        //公式计算器
        FormulaEvaluator formulaEvaluator = new HSSFFormulaEvaluator(sheets);

        //输出单元内容
        CellType cellType = cell.getCellType();
        System.out.println(cellType);
        switch (cellType) {
            case FORMULA:
                //得到公式
                String cellFormula = cell.getCellFormula();
                System.out.println(cellFormula);

                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String s = evaluate.formatAsString();
                System.out.println(s);

                break;
        }


    }


    */
/*======================写=============================*//*

    @Test
    public void bidDataWrite03() throws IOException {
        //记录开始时间
        long begin = System.currentTimeMillis();

        //创建一个SXSSFWorkbook
        Workbook workbook = new HSSFWorkbook();

        //创建一个sheet
        Sheet sheet = workbook.createSheet();

        //xls文件最大支持65536行
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            //创建一个行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {//创建单元格
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        System.out.println("done");
        FileOutputStream out = new FileOutputStream(PATH + "bigdata03.xls");
        workbook.write(out);
        // 操作结束，关闭文件
        out.close();

        //记录结束时间
        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin) / 1000);
    }

    */
/**
     * 大量数据测试
     *//*

    @Test
    public void bigDataWrite07() throws IOException {
        //记录开始时间
        long begin = System.currentTimeMillis();

        //创建一个SXSSFWorkbook
        Workbook workbook = new SXSSFWorkbook();

        //创建一个sheet
        Sheet sheet = workbook.createSheet();

        //xls文件最大支持65536行
        for (int rowNum = 0; rowNum < 100000; rowNum++) {
            //创建一个行
            Row row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {//创建单元格
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }

        System.out.println("done");
        FileOutputStream out = new FileOutputStream(PATH + "bigdata07-fast.xlsx");
        workbook.write(out);
        // 操作结束，关闭文件
        out.close();

        //清除临时文件
        ((SXSSFWorkbook) workbook).dispose();

        //记录结束时间
        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin) / 1000);
    }

    */
/**
     * 07 excel操作 -- 写
     *//*

    @Test
    public void write07() throws IOException {

        write07Method(15, 5);

    }

    public void write07Method(int rowNum, int cellNum) throws IOException {
        //创建新的excel工作簿  ， 对象变了
        XSSFWorkbook sheets = new XSSFWorkbook();

        //创建表
        XSSFSheet sheet = sheets.createSheet("test07sheet表");

        //行
        XSSFRow row = null;
        //单元格
        XSSFCell cell = null;

        //表头处理：
        row = sheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("单元格1-1");
        cell = row.createCell(1);
        cell.setCellValue("单元格1-2");
        cell = row.createCell(2);
        cell.setCellValue("单元格1-3");
        cell = row.createCell(3);
        cell.setCellValue("单元格1-4");
        cell = row.createCell(4);
        cell.setCellValue("单元格1-5");

        //数据
        for (int j = 1; j <= rowNum; j++) {
            //创建行
            row = sheet.createRow(j);
            for (int i = 0; i < cellNum; i++) {

                //创建单元格
                cell = row.createCell(i);
                cell.setCellValue(j + "--" + new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

            }
        }
        //输出流
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "粉丝统计.xlsx");

        //输出到工作簿
        sheets.write(fileOutputStream);

        //关闭
        fileOutputStream.close();

        System.out.println("生成成功！");
    }

    */
/**
     * 03 excel操作  —— 写
     *
     * @throws IOException
     *//*

    @Test
    public void write03() throws IOException {

        //创建新的Excel工作簿
        HSSFWorkbook sheets = new HSSFWorkbook();

        //在excel工作簿中建立一工作表，其中名称缺省值 sheet0
        HSSFSheet sheet = sheets.createSheet("test03sheet");

        //创建行 row  1
        HSSFRow row = sheet.createRow(0);

        //创建单元格 col 1-1
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("今日新增关注");

        //创建单元格 col 1-2
        HSSFCell cell1 = row.createCell(1);
        cell1.setCellValue(999);

        // 创建行（row 2）
        Row row2 = sheet.createRow(1);

        // 创建单元格（col 2-1）
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");

        //创建单元格
        Cell cell2 = row2.createCell(1);
        cell2.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //输出流
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "粉丝统计.xls");

        // 把相应的Excel 工作簿存盘
        sheets.write(fileOutputStream);

        // 操作结束，关闭文件
        fileOutputStream.close();

        System.out.println("文件生成成功");
    }
}
*/
