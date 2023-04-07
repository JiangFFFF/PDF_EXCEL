package to.jiangffff.test;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 *  高版本excel导出
 * @author JiangHuifeng
 * @create 2023-04-06-22:06
 */
public class POIDemo2 {

    public static void main(String[] args) throws IOException {
        // 创建工作簿
        Workbook workbook = new XSSFWorkbook();

        // 创建工作表
        Sheet sheet = workbook.createSheet("poi操作");

        // 创建行
        Row row = sheet.createRow(0);

        // 创建单元格
        Cell cell = row.createCell(0);
        cell.setCellValue("我是一个单元格");

        workbook.write(new FileOutputStream("/Users/jianghuifeng/Desktop/test.xlsx"));
    }

}
