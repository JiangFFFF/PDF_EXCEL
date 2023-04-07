package to.jiangffff.test;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * 读取excel
 * @author JiangHuifeng
 * @create 2023-04-06-22:25
 */
public class POIDemo3 {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("/Users/jianghuifeng/Desktop/基础教程/java报表/day1/资料/用户导入测试数据.xlsx"));
        // 获取工作表
        XSSFSheet sheet = workbook.getSheetAt(0);
        // 当前sheet最后一行的索引值
        int lastRowIndex = sheet.getLastRowNum();
        Row row =null;
        for (int i = 1; i <= lastRowIndex; i++) {
            row = sheet.getRow(i);
            String username = row.getCell(0).getStringCellValue();
            String phone = row.getCell(1).getStringCellValue();
            String province = row.getCell(2).getStringCellValue();
            String city = row.getCell(3).getStringCellValue();
            double numericCellValue = row.getCell(4).getNumericCellValue();
            String hireDate = row.getCell(5).getStringCellValue();
            String birthDay = row.getCell(6).getStringCellValue();
            String address = row.getCell(7).getStringCellValue();

        }

        // 读取内容

    }

}
