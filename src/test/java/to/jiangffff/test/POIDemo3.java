package to.jiangffff.test;

import com.itheima.pojo.User;
import com.itheima.service.UserService;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 读取excel
 * @author JiangHuifeng
 * @create 2023-04-06-22:25
 */
public class POIDemo3 {

    public static void main(String[] args) throws IOException, ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream("/Users/jianghuifeng/Desktop/基础教程/java报表/day1/资料/用户导入测试数据.xlsx"));
        // 获取工作表
        XSSFSheet sheet = workbook.getSheetAt(0);
        // 当前sheet最后一行的索引值
        int lastRowIndex = sheet.getLastRowNum();
        Row row =null;
        User user = null;
        // 读取内容
        for (int i = 1; i <= lastRowIndex; i++) {
            row = sheet.getRow(i);
            String username = row.getCell(0).getStringCellValue();
            String phone = null;
            try {
                phone = row.getCell(1).getStringCellValue();
            }catch (Exception e){
                phone = String.valueOf(row.getCell(1).getNumericCellValue());
            }

            String province = row.getCell(2).getStringCellValue();
            String city = row.getCell(3).getStringCellValue();
            Integer salary = ((Double)row.getCell(4).getNumericCellValue()).intValue();
            Date hireDate = row.getCell(5).getDateCellValue();
            Date birthDay = simpleDateFormat.parse(row.getCell(6).getStringCellValue());
            String address = row.getCell(7).getStringCellValue();
            user.setUserName(username);
            user.setPhone(phone);
            user.setProvince(province);
            user.setCity(city);
            user.setSalary(salary);
            user.setHireDate(hireDate);
            user.setBirthday(birthDay);
            user.setAddress(address);
        }



    }

}
