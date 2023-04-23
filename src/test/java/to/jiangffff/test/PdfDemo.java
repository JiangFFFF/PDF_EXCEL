package to.jiangffff.test;

import net.sf.jasperreports.engine.JREmptyDataSource;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * @author JiangHuifeng
 * @create 2023-04-16-14:57
 */
public class PdfDemo {
    public static void main(String[] args) throws Exception {
        // 模板文件
        String filePath = "/Users/jianghuifeng/Downloads/test01.jasper";
        FileInputStream inputStream = new FileInputStream(filePath);

        Map<String,Object> params = new HashMap<>();
        params.put("userNameP","张三");
        params.put("phoneP","15793736133");
        // 模板与数据的结合
        JasperPrint jasperPrint = JasperFillManager.fillReport(inputStream, params, new JREmptyDataSource());
        // 输出
        JasperExportManager.exportReportToPdfStream(jasperPrint,new FileOutputStream(new File("/Users/jianghuifeng/Downloads/test01.pdf")));
    }
}
