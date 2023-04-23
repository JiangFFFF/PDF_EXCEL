package com.itheima.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.StringUtils;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.Map;

/**
 * 自定义Excel引擎
 * @author JiangHuifeng
 * @create 2023-04-11-21:34
 */
public class ExcelExportEngine {

    public static Workbook writeToExcel(Object obj, Workbook workbook, String imagePath) throws IOException {
        Map<String, Object> map = EntityUtils.entityToMap(obj);
        Sheet sheet = workbook.getSheetAt(0);
        // 循环100行，每一行循环100个单元格
        Row row = null;
        Cell cell = null;
        for (int i = 0; i < 100; i++) {
            row = sheet.getRow(i);
            if(row == null){
                break;
            }else{
                for (int j = 0; j < 100; j++) {
                    cell = row.getCell(j);
                    if(cell != null){
                        writeToCell(cell,map);
                    }
                }
            }
        }
        if(imagePath != null){
            /**
             * 图片处理
             */
            // 创建输出流用于存储图片
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            // 读取图片，放入一个带有缓存区的图片类中
            BufferedImage bufferedImage = ImageIO.read(new File(imagePath));
            // 将图片写入字节输出流中
            String suffix = imagePath.substring(imagePath.lastIndexOf(".") + 1).toUpperCase();
            ImageIO.write(bufferedImage,suffix,byteArrayOutputStream);
            // Patriarch控制图片的写入；ClientAnchor指定图片的位置
            XSSFDrawing patriarch = (XSSFDrawing) sheet.createDrawingPatriarch();
            // 左上角x轴偏移 左上角y轴偏移  右下角x轴偏移 右下角y轴偏移  开始列 开始行 结束列 结束行
            // 偏移单位：是一个英式公制的单位 1厘米=360000；
            Sheet sheet1 = workbook.getSheetAt(1);
            int c0l1 = ((Double)sheet1.getRow(0).getCell(0).getNumericCellValue()).intValue();
            int row1 = ((Double)sheet1.getRow(0).getCell(1).getNumericCellValue()).intValue();
            int c0l2 = ((Double)sheet1.getRow(0).getCell(2).getNumericCellValue()).intValue();
            int row2 = ((Double)sheet1.getRow(0).getCell(3).getNumericCellValue()).intValue();
            XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, c0l1, row1, c0l2, row2);
            // 把图片写入sheet指定的位置
            int format = 0;
            switch (suffix){
                case "JPG":
                    format = XSSFWorkbook.PICTURE_TYPE_JPEG;
                    break;
                case "JPEG":
                    format = XSSFWorkbook.PICTURE_TYPE_JPEG;
                    break;
                case "PNG":
                    format = XSSFWorkbook.PICTURE_TYPE_PNG;
                    break;
                default:
            }
            patriarch.createPicture(anchor,workbook.addPicture(byteArrayOutputStream.toByteArray(), format));
            workbook.removeSheetAt(1);
        }
        return workbook;
    }

    /**
     * 比较单元格中的值是否在map中有匹配，如果有则将对应key的value放入单元格中
     * @param cell
     * @param map
     */
    private static void writeToCell(Cell cell, Map<String, Object> map) {
        CellType cellType = cell.getCellType();
        switch (cellType){
            case FORMULA:
                break;
            default:{
                String cellValue = cell.getStringCellValue();
                if(!StringUtils.isEmpty(cellValue)){
                    Map.Entry<String, Object> objectEntry = map.entrySet().stream().filter(item -> cellValue.equals(item.getKey())).findFirst().orElse(null);
                    if(objectEntry!=null){
                        cell.setCellValue(objectEntry.getValue().toString());
                    }
                }
            }
        }




    }

}
