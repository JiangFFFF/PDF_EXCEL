package to.jiangffff.test;

import com.itheima.pojo.User;
import com.opencsv.CSVReader;

import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;

/**
 *  读取百万数据的csv文件
 * @author JiangHuifeng
 * @create 2023-04-15-11:31
 */
public class CSVDemo {

    private static SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");

    public static void main(String[] args) throws IOException, ParseException {
        CSVReader csvReader = new CSVReader(new FileReader("/Users/jianghuifeng/Downloads/百万用户数据导出.csv"));
        // 小标题行
        String[] title = csvReader.readNext();
        User user = null;
        while (true){
            String[] content = csvReader.readNext();
            if(content == null){
                break;
            }
            user = new User();
            user.setId(Long.parseLong(content[0]));
            user.setUserName(content[1]);
            user.setPhone(content[2]);
            user.setHireDate(simpleDateFormat.parse(content[3]));
            user.setAddress(content[4]);
            System.out.println(user);
        }

    }
}
