package to.jiangffff.test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * 导出带样式的excel
 * @author JiangHuifeng
 * @create 2023-04-09-14:34
 */
public class POIDemo4 {
    public static void main(String[] args) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("有样式的excel");
        // 设置列宽
        sheet.setColumnWidth(0,256*10);
        sheet.setColumnWidth(1,256*15);
        sheet.setColumnWidth(2,256*15);
        sheet.setColumnWidth(3,256*15);
        sheet.setColumnWidth(4,256*30);
        Row bigTitleRow = sheet.createRow(0);
        // 设置行高
        bigTitleRow.setHeightInPoints(42);

        // 创建字体 黑体 18号
        Font font = workbook.createFont();
        font.setFontName("黑体");
        font.setFontHeightInPoints((short) 18);

        CellStyle bigTitleRowCellStyle = workbook.createCellStyle();
        // 设置上下左右边框
        // 细线
        bigTitleRowCellStyle.setBorderTop(BorderStyle.THIN);
        bigTitleRowCellStyle.setBorderBottom(BorderStyle.THIN);
        bigTitleRowCellStyle.setBorderLeft(BorderStyle.MEDIUM);
        bigTitleRowCellStyle.setBorderRight(BorderStyle.MEDIUM);
        // 设置字体
        bigTitleRowCellStyle.setFont(font);
        // 对齐方式
        // 水平居中对齐
        bigTitleRowCellStyle.setAlignment(HorizontalAlignment.CENTER);
        // 垂直居中对齐
        bigTitleRowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        for (int i = 0; i < 5; i++) {
            Cell cell = bigTitleRow.createCell(i);
            cell.setCellStyle(bigTitleRowCellStyle);
        }
        // 合并单元格
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,4));
        // 放入数据
        sheet.getRow(0).getCell(0).setCellValue("用户信息数据");

        // 小标题样式
        CellStyle littleTitleRowCellStyle = workbook.createCellStyle();
        // 样式克隆
        littleTitleRowCellStyle.cloneStyleFrom(bigTitleRowCellStyle);
        // 设置字体 宋体 12号 加粗
        Font littleFont = workbook.createFont();
        littleFont.setFontName("宋体");
        littleFont.setFontHeightInPoints((short)12);
        littleFont.setBold(true);
        littleTitleRowCellStyle.setFont(littleFont);

        // 内容样式
        CellStyle contentRowCellStyle = workbook.createCellStyle();
        // 样式克隆
        contentRowCellStyle.cloneStyleFrom(bigTitleRowCellStyle);
        contentRowCellStyle.setAlignment(HorizontalAlignment.LEFT);
        // 设置字体 宋体 11号
        Font contentFont = workbook.createFont();
        contentFont.setFontName("宋体");
        contentFont.setFontHeightInPoints((short)11);
        contentRowCellStyle.setFont(contentFont);

        Row titelRow = sheet.createRow(1);
        titelRow.setHeightInPoints(32);
        String[] titles = new String[]{"编号","姓名","手机号","入职日期","现住址"};
        for (int i = 0; i < titles.length; i++) {
            Cell cell = titelRow.createCell(i);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(littleTitleRowCellStyle);
        }

        Row contentRow = sheet.createRow(2);
        String[] content = new String[]{"1","大一","135555555","2011-01-01","住址"};
        for (int i = 0; i < content.length; i++) {
            Cell cell = titelRow.createCell(i);
            cell.setCellValue(content[i]);
            cell.setCellStyle(contentRowCellStyle);
        }


        workbook.write(new FileOutputStream("/Users/jianghuifeng/Desktop/testStyle.xlsx"));
    }
}

