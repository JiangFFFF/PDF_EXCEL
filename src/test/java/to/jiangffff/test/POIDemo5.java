package to.jiangffff.test;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

/**
 *  测试百万数据导入
 * @author JiangHuifeng
 * @create 2023-04-13-21:39
 */
public class POIDemo5 {
    public static void main(String[] args) throws Exception {
//        XSSFWorkbook workbook = new XSSFWorkbook("/Users/jianghuifeng/Desktop/百万用户数据导出.xlsx");
//        XSSFSheet sheetAt = workbook.getSheetAt(0);
//        String stringCellValue = sheetAt.getRow(0).getCell(0).getStringCellValue();
//        System.out.println(stringCellValue);

        new ExcelParse().parse("/Users/jianghuifeng/Desktop/百万用户数据导出.xlsx");

    }
}
