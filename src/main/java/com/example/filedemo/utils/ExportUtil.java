package com.example.filedemo.utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.List;

public class ExportUtil {

    /**
     * 导出Excel
     *
     * @param sheetName sheet名称
     * @param title     标题
     * @param values    内容，可以修改为list  或者 map
     * @param wb        HSSFWorkbook对象
     * @return
     */
    public static HSSFWorkbook getHSSFWorkbook(String sheetName, String[] title, String[][] values, HSSFWorkbook wb) {

        // 第一步，创建一个HSSFWorkbook，对应一个Excel文件
        if (wb == null) {
            wb = new HSSFWorkbook();
        }

        // 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet(sheetName);

        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制
        HSSFRow row = sheet.createRow(0);

        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER); // 创建一个居中格式

        //声明列对象
        HSSFCell cell = null;

        //创建标题
        for (int i = 0; i < title.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
            cell.setCellStyle(style);
        }
        //创建内容
        for (int i = 0; i < values.length; i++) {
            row = sheet.createRow(i + 1);
            for (int j = 0; j < values[i].length; j++) {
                //将内容按顺序赋给对应的列对象
                row.createCell(j).setCellValue(values[i][j]);
            }
        }
        return wb;
    }

    /**
     * 转成excel文件
     * 调用该方法一定要在外层关闭流
     *
     * @return
     */
    public static XSSFWorkbook getXSSFWorkbook(String sheetName, List<String> headList, List<List> contentList) throws IOException {

        XSSFWorkbook wb = new XSSFWorkbook();

        XSSFSheet sheet = wb.createSheet(sheetName);
        //设置字体
        Font font = wb.createFont();
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 11);

        //设置单元格样式
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setFont(font);

        int rowPos = 0;
        Row headRow = sheet.createRow(rowPos++);

        for (int i = 0; i < headList.size(); i++) {
            sheet.setDefaultColumnStyle(0, cellStyle);
            headRow.createCell(i).setCellValue(headList.get(i));
        }

        for (int i = 0; i < contentList.size(); i++) {

            List rowContent = contentList.get(i);
            Row row = sheet.createRow(rowPos++);

            for (int j = 0; j < rowContent.size(); j++) {
                Object cellContent = rowContent.get(j);
                Cell cell = row.createCell(j);
                if (cellContent == null) {
                    cell.setCellValue("");
                } else {

                    if (cellContent instanceof BigDecimal) {
                        cell.setCellValue(((BigDecimal) cellContent).intValue());
                    } else {
                        cell.setCellValue(cellContent.toString());
                    }
                }
            }
        }

        ByteArrayOutputStream content = new ByteArrayOutputStream();
        wb.write(content);

        return wb;
    }
}

