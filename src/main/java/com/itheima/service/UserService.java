package com.itheima.service;

import cn.afterturn.easypoi.csv.CsvExportUtil;
import cn.afterturn.easypoi.csv.entity.CsvExportParams;
import cn.afterturn.easypoi.entity.ImageEntity;
import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.word.WordExportUtil;
import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.itheima.mapper.ResourceMapper;
import com.itheima.mapper.UserMapper;
import com.itheima.pojo.Resource;
import com.itheima.pojo.User;
//import jxl.Workbook;
//import org.apache.poi.ss.usermodel.Workbook;
import com.itheima.utils.EntityUtils;
import com.itheima.utils.ExcelExportEngine;
import com.opencsv.CSVWriter;
import com.zaxxer.hikari.HikariDataSource;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import net.sf.jasperreports.engine.JREmptyDataSource;
import net.sf.jasperreports.engine.JasperExportManager;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ResourceUtils;
import org.springframework.web.multipart.MultipartFile;
import tk.mybatis.mapper.entity.Example;


import javax.imageio.ImageIO;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

@Service
public class UserService {

    private SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");


    @Autowired
    private UserMapper userMapper;

    @Autowired
    private ResourceMapper resourceMapper;

    public List<User> findAll() {
        return userMapper.selectAll();
    }

    public List<User> findPage(Integer page, Integer pageSize) {
        PageHelper.startPage(page, pageSize);  //开启分页
        Page<User> userPage = (Page<User>) userMapper.selectAll(); //实现查询
        return userPage.getResult();
    }

    public User findById(Long id) {
        User user = userMapper.selectByPrimaryKey(id);
        // 查询办公用品
        Resource resource = new Resource();
        resource.setUserId(id);
        List<Resource> resourceList = resourceMapper.select(resource);
        user.setResourceList(resourceList);
        return user;
    }

    public void downLoadXlsByJxl(HttpServletResponse response) throws IOException, WriteException {
        ServletOutputStream outputStream = response.getOutputStream();
        WritableWorkbook workbook = Workbook.createWorkbook(outputStream);
        WritableSheet sheet = workbook.createSheet("jxl入门", 0);
        // 第一个参数 列的索引，第二个参数 标准字母的宽度
        sheet.setColumnView(0, 5);
        sheet.setColumnView(1, 10);
        sheet.setColumnView(2, 20);
        sheet.setColumnView(3, 5);
        sheet.setColumnView(4, 5);
        String[] title = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        for (int i = 0; i < title.length; i++) {
            Label label = new Label(i, 0, title[i]);
            sheet.addCell(label);
        }
        List<User> userList = userMapper.selectAll();
        int count = 0;
        for (User user : userList) {
            count++;
            Label label1 = new Label(0, count, String.valueOf(user.getId()));
            sheet.addCell(label1);
            Label label2 = new Label(1, count, user.getUserName());
            sheet.addCell(label2);
            Label label3 = new Label(2, count, user.getPhone());
            sheet.addCell(label3);
            Date hireDate = user.getHireDate();
            Label label4 = new Label(3, count, simpleDateFormat.format(hireDate));
            sheet.addCell(label4);
            Label label5 = new Label(4, count, user.getAddress());
            sheet.addCell(label5);
        }

        // 导出文件 一个流（outputStream）两个头(文件打开方式 in-line attachment,文件下载时mime类型 application/vnd.ms-excel)
        String filename = "一个jmx入门.xls";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("application/vnd.ms-excel");
        workbook.write();
        workbook.close();
        outputStream.close();
    }

    public void uploadExcel(MultipartFile file) throws IOException, ParseException {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
        // 获取工作表
        XSSFSheet sheet = workbook.getSheetAt(0);
        // 当前sheet最后一行的索引值
        int lastRowIndex = sheet.getLastRowNum();
        Row row = null;
        User user = null;
        // 读取内容
        for (int i = 1; i <= lastRowIndex; i++) {
            row = sheet.getRow(i);
            String username = row.getCell(0).getStringCellValue();
            String phone = null;
            try {
                phone = row.getCell(1).getStringCellValue();
            } catch (Exception e) {
                phone = String.valueOf(row.getCell(1).getNumericCellValue());
            }

            String province = row.getCell(2).getStringCellValue();
            String city = row.getCell(3).getStringCellValue();
            Integer salary = ((Double) row.getCell(4).getNumericCellValue()).intValue();
            Date hireDate = simpleDateFormat.parse(row.getCell(5).getStringCellValue());
            Date birthDay = simpleDateFormat.parse(row.getCell(6).getStringCellValue());
            String address = row.getCell(7).getStringCellValue();
            user = new User();
            user.setUserName(username);
            user.setPhone(phone);
            user.setProvince(province);
            user.setCity(city);
            user.setSalary(salary);
            user.setHireDate(hireDate);
            user.setBirthday(birthDay);
            user.setAddress(address);
            userMapper.insert(user);
        }
    }

    /**
     * 不带样式导出
     *
     * @param response
     * @throws IOException
     */
    public void downLoadXlsxByPoi(HttpServletResponse response) throws IOException {
        // 1、创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 2、创建工作表
        Sheet sheet = workbook.createSheet("用户数据");
        // 设置列宽 1代表一个标准字母宽度的256分之一
        sheet.setColumnWidth(2, 15 * 256);
        sheet.setColumnWidth(3, 15 * 256);
        sheet.setColumnWidth(4, 35 * 256);
        // 3、处理固定标题
        String[] title = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        Row titleRow = sheet.createRow(0);
        Cell cell = null;
        ;
        for (int i = 0; i < title.length; i++) {
            cell = titleRow.createCell(i);
            cell.setCellValue(title[i]);
        }
        // 4、从第二行循环遍历数据
        List<User> userList = userMapper.selectAll();
        int rowIndex = 1;
        Row row = null;
        for (User user : userList) {
            row = sheet.createRow(rowIndex);
            cell = row.createCell(0);
            cell.setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getUserName());
            row.createCell(2).setCellValue(user.getPhone());
            row.createCell(3).setCellValue(simpleDateFormat.format(user.getHireDate()));
            row.createCell(4).setCellValue(user.getAddress());
            rowIndex++;
        }
        // 一个流两个头
        String filename = "员工数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    /**
     * 带样式导出
     *
     * @param response
     * @throws IOException
     */
    public void downLoadXlsxByPoiWithCellStyle(HttpServletResponse response) throws IOException {
        org.apache.poi.ss.usermodel.Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("有样式的excel");
        // 设置列宽
        sheet.setColumnWidth(0, 256 * 10);
        sheet.setColumnWidth(1, 256 * 15);
        sheet.setColumnWidth(2, 256 * 15);
        sheet.setColumnWidth(3, 256 * 15);
        sheet.setColumnWidth(4, 256 * 30);
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
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));
        // 放入数据
        sheet.getRow(0).getCell(0).setCellValue("用户信息数据");

        // 小标题样式
        CellStyle littleTitleRowCellStyle = workbook.createCellStyle();
        // 样式克隆
        littleTitleRowCellStyle.cloneStyleFrom(bigTitleRowCellStyle);
        // 设置字体 宋体 12号 加粗
        Font littleFont = workbook.createFont();
        littleFont.setFontName("宋体");
        littleFont.setFontHeightInPoints((short) 12);
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
        contentFont.setFontHeightInPoints((short) 11);
        contentRowCellStyle.setFont(contentFont);

        Row titelRow = sheet.createRow(1);
        titelRow.setHeightInPoints(32);
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        for (int i = 0; i < titles.length; i++) {
            Cell cell = titelRow.createCell(i);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(littleTitleRowCellStyle);
        }

        List<User> userList = userMapper.selectAll();
        int rowIndex = 2;
        Row row = null;
        Cell cell = null;
        for (User user : userList) {
            row = sheet.createRow(rowIndex);
            cell = row.createCell(0);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getId());

            cell = row.createCell(1);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getUserName());

            cell = row.createCell(2);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getPhone());

            cell = row.createCell(3);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(simpleDateFormat.format(user.getHireDate()));

            cell = row.createCell(4);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getAddress());
            rowIndex++;
        }

//        workbook.write(new FileOutputStream("/Users/jianghuifeng/Desktop/testStyle.xlsx"));
        // 一个流两个头
        String filename = "员工数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    /**
     * 模板导出
     *
     * @param response
     */
    public void downLoadXlsxByPoiWithTemplate(HttpServletResponse response) throws IOException, InvalidFormatException {
        // 1、获取模板
        // 项目根目录
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile, "/excel_template/userList.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(templateFile);
        XSSFSheet sheet = workbook.getSheetAt(0);
        // 获取准备好的单元格样式
        XSSFCellStyle contentRowCellStyle = workbook.getSheetAt(1).getRow(0).getCell(0).getCellStyle();
        // 2、查询员工数据
        List<User> userList = userMapper.selectAll();
        // 3、将数据放入模板
        int rowIndex = 2;
        Row row = null;
        Cell cell = null;
        for (User user : userList) {
            row = sheet.createRow(rowIndex);
            row.setHeightInPoints(15);
            cell = row.createCell(0);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getId());

            cell = row.createCell(1);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getUserName());

            cell = row.createCell(2);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getPhone());

            cell = row.createCell(3);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(simpleDateFormat.format(user.getHireDate()));

            cell = row.createCell(4);
            cell.setCellStyle(contentRowCellStyle);
            cell.setCellValue(user.getAddress());
            rowIndex++;
        }

        // 删除第二个sheet
        workbook.removeSheetAt(1);

        // 4、导出文件
        String filename = "员工数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    /**
     * 使用模板导出用户详细数据
     *
     * @param id
     * @param response
     */
    public void downloadUserInfoByTemplate(Long id, HttpServletResponse response) throws IOException, InvalidFormatException {
        // 1、读取模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile, "/excel_template/userInfo.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(templateFile);
        XSSFSheet sheet = workbook.getSheetAt(0);
        // 2、根据id获取用户信息
        User user = userMapper.selectByPrimaryKey(id);
        // 3、数据放入模板
        // 用户名 第2行第2列
        sheet.getRow(1).getCell(1).setCellValue(user.getUserName());
        // 手机号 第3行第2列
        sheet.getRow(2).getCell(1).setCellValue(user.getPhone());
        // 生日   第4行第2列
        sheet.getRow(3).getCell(1).setCellValue(simpleDateFormat.format(user.getBirthday()));
        // 工资   第5行第2列
        sheet.getRow(4).getCell(1).setCellValue(user.getSalary());
        // 入职日期 第6行第2列
        sheet.getRow(5).getCell(1).setCellValue(simpleDateFormat.format(user.getHireDate()));
        // 省份   第7行第2列
        sheet.getRow(6).getCell(1).setCellValue(user.getProvince());
        // 现住址  第8行第2列
        sheet.getRow(7).getCell(1).setCellValue(user.getAddress());
        // 司龄    第6行第4列
        /**
         * 公式处理司龄
         */
        sheet.getRow(5).getCell(3).setCellFormula("CONCATENATE(DATEDIF(B6,TODAY(),\"Y\"),\"年\",DATEDIF(B6,TODAY(),\"YM\"),\"个月\")");


        // 城市    第7行第4列
        sheet.getRow(6).getCell(3).setCellValue(user.getCity());
        // 照片    第2行至第5行，第3列至第4列
        /**
         * 图片处理
         */
        // 创建输出流用于存储图片
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        // 读取图片，放入一个带有缓存区的图片类中
        BufferedImage bufferedImage = ImageIO.read(new File(rootFile, user.getPhoto()));
        // 将图片写入字节输出流中
        String suffix = user.getPhoto().substring(user.getPhoto().lastIndexOf(".") + 1).toUpperCase();
        ImageIO.write(bufferedImage, suffix, byteArrayOutputStream);
        // Patriarch控制图片的写入；ClientAnchor指定图片的位置
        XSSFDrawing patriarch = sheet.createDrawingPatriarch();
        // 左上角x轴偏移 左上角y轴偏移  右下角x轴偏移 右下角y轴偏移  开始列 开始行 结束列 结束行
        // 偏移单位：是一个英式公制的单位 1厘米=360000；
        XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, 0, 0, 2, 1, 4, 5);
        // 把图片写入sheet指定的位置
        int format = 0;
        switch (suffix) {
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
        patriarch.createPicture(anchor, workbook.addPicture(byteArrayOutputStream.toByteArray(), format));

        // 4、导出
        String filename = "员工(" + user.getUserName() + ")详细信息.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    public void downloadUserInfoByTemplate2(Long id, HttpServletResponse response) throws IOException, InvalidFormatException {
        // 1、读取模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile, "/excel_template/userInfo2.xlsx");
        // 2、根据id获取用户信息
        User user = userMapper.selectByPrimaryKey(id);
        org.apache.poi.ss.usermodel.Workbook workbook = new XSSFWorkbook(templateFile);
        workbook = ExcelExportEngine.writeToExcel(user, workbook, rootFile.getPath() + user.getPhoto());
        // 3、导出
        String filename = "员工(" + user.getUserName() + ")详细信息.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    /**
     * 导出百万数据
     * 1、使用该版本的exel
     * 2、使用sax方式解析Excel（解析XML）
     * 3、不能使用模板
     * 4、不能使用太多样式
     *
     * @param response
     */
    public void downLoadMillion(HttpServletResponse response) throws IOException {
        // 指定sax方式
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        int page = 1;
        // 记录处理数据的个数
        int num = 0;
        // 每个sheet的行索引
        int rowIndex = 1;
        Row row = null;
        SXSSFSheet sheet = null;
        while (true) {
            List<User> userList = this.findPage(page, 10000);
            if (CollectionUtils.isEmpty(userList)) {
                break;
            }
            if (num % 1000000 == 0) {
                sheet = workbook.createSheet("第" + ((num / 1000000) + 1) + "个工作表");
                // 每个新sheet中的行索引重置
                rowIndex = 1;
                // 设置小标题
                String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
                SXSSFRow titleRow = sheet.createRow(0);
                for (int i = 0; i < titles.length; i++) {
                    titleRow.createCell(i).setCellValue(titles[i]);
                }
            }
            for (User user : userList) {
                row = sheet.createRow(rowIndex);
                row.createCell(0).setCellValue(user.getId());
                row.createCell(1).setCellValue(user.getUserName());
                row.createCell(2).setCellValue(user.getPhone());
                row.createCell(3).setCellValue(simpleDateFormat.format(user.getHireDate()));
                row.createCell(4).setCellValue(user.getAddress());
                rowIndex++;
                num++;
            }
            page++;
        }

        String filename = "百万用户数据导出.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        workbook.write(response.getOutputStream());
        workbook.close();

    }

    /**
     * 用csv文件导出百万数据
     *
     * @param response
     */
    public void downLoadCSV(HttpServletResponse response) throws IOException {
        ServletOutputStream outputStream = response.getOutputStream();
        CSVWriter csvWriter = new CSVWriter(new OutputStreamWriter(outputStream, "UTF-8"));
        String filename = "百万用户数据导出.csv";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("text/csv");
        String[] titles = new String[]{"编号", "姓名", "手机号", "入职日期", "现住址"};
        // 写入小标题数据
        csvWriter.writeNext(titles);
        int page = 1;
        while (true) {
            List<User> userList = this.findPage(page, 200000);
            if (CollectionUtils.isEmpty(userList)) {
                break;
            }
            for (User user : userList) {
                csvWriter.writeNext(
                        new String[]{user.getId().toString(),
                                user.getUserName(),
                                user.getPhone(),
                                simpleDateFormat.format(user.getHireDate()),
                                user.getAddress()});
            }
            page++;
            csvWriter.flush();
        }
        csvWriter.close();

    }


    /**
     * 下载用户合同文档
     * @param id
     * @param response
     */
    public void downloadContract(Long id, HttpServletResponse response) throws IOException {
        // 读取到模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile, "/word_template/contract_template.docx");
        XWPFDocument word = new XWPFDocument(new FileInputStream(templateFile));
        // 查询当前用户数据
        User user = this.findById(id);
        Map<String,Object> params = new HashMap<>();
        params.put("userName",user.getUserName());
        params.put("hireDate",simpleDateFormat.format(user.getHireDate()));
        params.put("address",user.getAddress());
        // 替换模板中的数据
        // 处理正文
        List<XWPFParagraph> paragraphs = word.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                String text = run.getText(0);
                Map.Entry<String, Object> entry = params.entrySet().stream().filter(item -> text.contains(item.getKey())).findFirst().orElse(null);
                if(entry != null){
                    run.setText(text.replaceAll(entry.getKey(),entry.getValue().toString()),0);
                }
            }
        }
        // 处理表格
        List<Resource> resourceList = user.getResourceList();
        XWPFTable xwpfTable = word.getTables().get(0);
        XWPFTableRow row = xwpfTable.getRow(0);
        int rowIndex = 1;
        for (Resource resource : resourceList) {
            // 添加行
//            xwpfTable.addRow(row);
            // 拷贝行
            copyRow(xwpfTable,row,rowIndex);
            XWPFTableRow row1 = xwpfTable.getRow(rowIndex);
            row1.getCell(0).setText(resource.getName());
            row1.getCell(1).setText(resource.getPrice().toString());
            row1.getCell(2).setText(resource.getNeedReturn()?"需要":"不需要");
            File imageFile = new File(rootFile+"/static"+resource.getPhoto());
            setCellImage(row1.getCell(3),imageFile);
            rowIndex++;

        }



        // 导出word
        String filename = "员工(" + user.getUserName() + ")合同.docx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        word.write(response.getOutputStream());
        word.close();


    }


    /**
     * 用于深克隆行
     * @param xwpfTable
     * @param sourceRow 被复制的行
     * @param rowIndex 复制后的行标
     */
    private void copyRow(XWPFTable xwpfTable, XWPFTableRow sourceRow, int rowIndex) {
        XWPFTableRow targetRow = xwpfTable.insertNewTableRow(rowIndex);
        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
        // 获取源行的单元格
        List<XWPFTableCell> cells = sourceRow.getTableCells();
        if(CollectionUtils.isEmpty(cells)){
           return;
        }
        XWPFTableCell targetCell = null;
        for (XWPFTableCell cell : cells) {
            targetCell = targetRow.addNewTableCell();
            // 复制单元格属性
            targetCell.getCTTc().setTcPr(cell.getCTTc().getTcPr());
            // 复制段落属性
            targetCell.getParagraphs().get(0).getCTP().setPPr(cell.getParagraphs().get(0).getCTP().getPPr());
        }
    }

    /**
     * 向单元格中写入图片
     * @param cell
     * @param imageFile
     */
    private void setCellImage(XWPFTableCell cell, File imageFile) {
        XWPFRun run = cell.getParagraphs().get(0).createRun();
        try (FileInputStream fileInputStream = new FileInputStream(imageFile)){
            run.addPicture(fileInputStream,XWPFDocument.PICTURE_TYPE_JPEG,imageFile.getName(), Units.toEMU(100),Units.toEMU(50));
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 用easyPOI方式导出excel
     * @param response
     */
    public void downLoadWithEasyPOI(HttpServletResponse response) throws IOException {
        ExportParams exportParams = new ExportParams("员工信息列表数据","数据sheet", ExcelType.XSSF);
        List<User> userList = userMapper.selectAll();
        org.apache.poi.ss.usermodel.Workbook workbook = ExcelExportUtil.exportExcel(exportParams, User.class, userList);
        String filename = "用户数据导出.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(response.getOutputStream());
    }

    /**
     * 用easyPOI方式导入excel
     * @param file
     */
    public void uploadExcelWithEasyPOI(MultipartFile file) throws Exception {
        ImportParams importParams = new ImportParams();
        importParams.setNeedSave(false);
        importParams.setHeadRows(1);
        importParams.setHeadRows(1);
        List<User> userList = ExcelImportUtil.importExcel(file.getInputStream(), User.class, importParams);
        for (User user : userList) {
            user.setId(null);
            userMapper.insert(user);
        }
    }

    /**
     * 用easyPOI方式利用模板导出excel
     * @param id
     * @param response
     */
    public void downloadUserInfoByEasyPOI(Long id, HttpServletResponse response) throws IOException {
        // 1、读取模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile, "/excel_template/userInfo3.xlsx");
        // 2、根据id获取用户信息
        User user = userMapper.selectByPrimaryKey(id);
        Map<String, Object> map = EntityUtils.entityToMap(user);
        ImageEntity imageEntity = new ImageEntity();
        imageEntity.setUrl(user.getPhone());
        // 占用多少列
        imageEntity.setColspan(2);
        // 占用多少行
        imageEntity.setRowspan(4);
        map.put("photo",imageEntity);

        TemplateExportParams exportParams = new TemplateExportParams(templateFile.getPath(),true);
        org.apache.poi.ss.usermodel.Workbook workbook = ExcelExportUtil.exportExcel(exportParams, map);
        String filename = "用户数据.xlsx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        workbook.write(response.getOutputStream());
    }

    /**
     * 利用easyPOI导出csv文件
     * @param response
     */
    public void downLoadCSVWithEasyPOI(HttpServletResponse response) throws IOException {
        String filename = "用户数据导出.csv";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("text/csv");
        CsvExportParams csvExportParams = new CsvExportParams();
        // 忽略照片一列
        csvExportParams.setExclusions(new String[]{"照片"});
        List<User> userList = userMapper.selectAll();
        CsvExportUtil.exportCsv(csvExportParams,User.class,userList,response.getOutputStream());

    }

    /**
     * 利用easyPOI导出用户的合同文档
     * @param id
     * @param response
     */
    public void downloadContractByEasyPOI(Long id, HttpServletResponse response) throws Exception {
        // 读取到模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile, "/word_template/contract_template2.docx");
        User user = this.findById(id);
        Map<String,Object> params = new HashMap<>();
        params.put("userName",user.getUserName());
        params.put("hireDate",simpleDateFormat.format(user.getHireDate()));
        params.put("address",user.getAddress());

        // 测试
        ImageEntity imageEntityContent = new ImageEntity();
        imageEntityContent.setUrl(rootFile.getPath()+user.getPhoto());
        imageEntityContent.setWidth(100);
        imageEntityContent.setHeight(50);
        params.put("photo",imageEntityContent);

        List<Map<String,Object>> resourceMapList = new ArrayList<>();
        Map<String,Object> map = null;
        for (Resource resource : user.getResourceList()) {
            map = new HashMap<>();
            map.put("name",resource.getName());
            map.put("price",resource.getPrice());
            map.put("needReturn",resource.getNeedReturn());

            // 处理照片
            ImageEntity imageEntity = new ImageEntity();
            imageEntity.setUrl(rootFile.getPath()+"/static"+resource.getPhoto());
            map.put("photo",imageEntity);

            resourceMapList.add(map);
        }
        params.put("resourceList",resourceMapList);

        XWPFDocument word = WordExportUtil.exportWord07(templateFile.getPath(), params);
        // 导出word
        String filename = "员工(" + user.getUserName() + ")合同.docx";
        response.setHeader("content-disposition", "attachment;filename=" + new String(filename.getBytes(), "ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        word.write(response.getOutputStream());
        word.close();
    }

    @Autowired
    private HikariDataSource hikariDataSource;

    /**
     * 导出用户数据到pdf中（直接从数据库中导出）
     * @param response
     */
    public void downLoadPDF(HttpServletResponse response) throws Exception {
        // 获取模板文件
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile, "/pdf_template/userList_db.jasper");
        Map<String,Object> map = new HashMap<>();
//        JasperPrint jasperPrint = JasperFillManager.fillReport(new FileInputStream(templateFile), map, getCon());
        JasperPrint jasperPrint = JasperFillManager.fillReport(new FileInputStream(templateFile), map, hikariDataSource.getConnection());
        ServletOutputStream outputStream = response.getOutputStream();
        String filename="用户列表数据.pdf";
        response.setContentType("application/pdf");
        response.setHeader("content-disposition","attachment;filename="+new String(filename.getBytes(),"ISO8859-1"));
        JasperExportManager.exportReportToPdfStream(jasperPrint,outputStream);
    }

    private Connection getCon() throws ClassNotFoundException, SQLException {
        Class.forName("com.mysql.jdbc.Driver");
        Connection connection = DriverManager.getConnection("jdbc:mysql://192.168.31.32:3306/report_manager_db?useSSL=false&useUnicode=true&characterEncoding=utf-8&zeroDateTimeBehavior=convertToNull&transformedBitIsBoolean=true&serverTimezone=GMT%2B8&nullCatalogMeansCurrent=true&allowPublicKeyRetrieval=true", "root", "19990926");
        return connection;
    }

    /**
     * 导出用户数据到pdf中（后台导出）
     * @param response
     * @throws Exception
     */
    public void downLoadPDF2(HttpServletResponse response) throws Exception {
        // 获取模板文件
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile, "/pdf_template/userList2.jasper");


        Map<String,Object> params = new HashMap<>();
        Example example = new Example(User.class);
        example.setOrderByClause("province");
        List<User> userList = userMapper.selectByExample(example);
        userList = userList.stream().map(item->{
                item.setHireDateStr(simpleDateFormat.format(item.getHireDate()));
                return item;
        }).collect(Collectors.toList());
        JRBeanCollectionDataSource dataSource = new JRBeanCollectionDataSource(userList);

        JasperPrint jasperPrint = JasperFillManager.fillReport(new FileInputStream(templateFile), params,dataSource);
        ServletOutputStream outputStream = response.getOutputStream();


        String filename="用户列表数据.pdf";
        response.setContentType("application/pdf");
        response.setHeader("content-disposition","attachment;filename="+new String(filename.getBytes(),"ISO8859-1"));
        JasperExportManager.exportReportToPdfStream(jasperPrint,outputStream);
    }

    /**
     * 导出用户详细信息pdf
     * @param id
     * @param response
     */
    public void downloadUserInfoByPDF(Long id, HttpServletResponse response) throws Exception {
        // 获取模板文件
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile, "/pdf_template/userInfo.jasper");

        User user = userMapper.selectByPrimaryKey(id);
        Map<String, Object> params = EntityUtils.entityToMap(user);
        params.put("salary",user.getSalary().toString());
        params.put("photo",rootFile.getPath()+user.getPhoto());

        JasperPrint jasperPrint = JasperFillManager.fillReport(new FileInputStream(templateFile), params,new JREmptyDataSource());
        ServletOutputStream outputStream = response.getOutputStream();


        String filename="用户详细数据.pdf";
        response.setContentType("application/pdf");
        response.setHeader("content-disposition","attachment;filename="+new String(filename.getBytes(),"ISO8859-1"));
        JasperExportManager.exportReportToPdfStream(jasperPrint,outputStream);
    }
}
