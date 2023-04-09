package com.itheima.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.itheima.mapper.UserMapper;
import com.itheima.pojo.User;
//import jxl.Workbook;
//import org.apache.poi.ss.usermodel.Workbook;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.ResourceUtils;
import org.springframework.web.multipart.MultipartFile;


import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

@Service
public class UserService {

    private SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");


    @Autowired
    private UserMapper userMapper;

    public List<User> findAll() {
        return userMapper.selectAll();
    }

    public List<User> findPage(Integer page, Integer pageSize) {
        PageHelper.startPage(page,pageSize);  //开启分页
        Page<User> userPage = (Page<User>) userMapper.selectAll(); //实现查询
        return userPage.getResult();
    }

    public void downLoadXlsByJxl(HttpServletResponse response) throws IOException, WriteException {
        ServletOutputStream outputStream = response.getOutputStream();
        WritableWorkbook workbook = Workbook.createWorkbook(outputStream);
        WritableSheet sheet = workbook.createSheet("jxl入门", 0);
        // 第一个参数 列的索引，第二个参数 标准字母的宽度
        sheet.setColumnView(0,5);
        sheet.setColumnView(1,10);
        sheet.setColumnView(2,20);
        sheet.setColumnView(3,5);
        sheet.setColumnView(4,5);
        String[] title = new String[]{"编号","姓名","手机号","入职日期","现住址"};
        for (int i = 0; i < title.length; i++) {
            Label label = new Label(i, 0, title[i]);
            sheet.addCell(label);
        }
        List<User> userList = userMapper.selectAll();
        int count =0;
        for (User user : userList) {
            count ++;
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
        String filename="一个jmx入门.xls";
        response.setHeader("content-disposition","attachment;filename="+new String(filename.getBytes(),"ISO-8859-1"));
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
     * @param response
     * @throws IOException
     */
    public void downLoadXlsxByPoi(HttpServletResponse response) throws IOException {
        // 1、创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 2、创建工作表
        Sheet sheet = workbook.createSheet("用户数据");
        // 设置列宽 1代表一个标准字母宽度的256分之一
        sheet.setColumnWidth(2,15*256);
        sheet.setColumnWidth(3,15*256);
        sheet.setColumnWidth(4,35*256);
        // 3、处理固定标题
        String[] title = new String[]{"编号","姓名","手机号","入职日期","现住址"};
        Row titleRow = sheet.createRow(0);
        Cell cell = null;;
        for (int i = 0; i < title.length; i++) {
            cell = titleRow.createCell(i);
            cell.setCellValue(title[i]);
        }
        // 4、从第二行循环遍历数据
        List<User> userList = userMapper.selectAll();
        int rowIndex =1;
        Row row = null;
        for (User user : userList) {
            row = sheet.createRow(rowIndex);
            cell = row.createCell(0);
            cell.setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getUserName());
            row.createCell(2).setCellValue(user.getPhone());
            row.createCell(3).setCellValue(simpleDateFormat.format(user.getHireDate()));
            row.createCell(4).setCellValue(user.getAddress());
            rowIndex ++;
        }
        // 一个流两个头
        String filename="员工数据.xlsx";
        response.setHeader("content-disposition","attachment;filename=" + new String(filename.getBytes(),"ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    /**
     * 带样式导出
     * @param response
     * @throws IOException
     */
    public void downLoadXlsxByPoiWithCellStyle(HttpServletResponse response) throws IOException {
        org.apache.poi.ss.usermodel.Workbook workbook = new XSSFWorkbook();
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
            rowIndex ++;
        }

//        workbook.write(new FileOutputStream("/Users/jianghuifeng/Desktop/testStyle.xlsx"));
        // 一个流两个头
        String filename="员工数据.xlsx";
        response.setHeader("content-disposition","attachment;filename=" + new String(filename.getBytes(),"ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    /**
     * 模板导出
     * @param response
     */
    public void downLoadXlsxByPoiWithTemplate(HttpServletResponse response) throws IOException, InvalidFormatException {
        // 1、获取模板
        // 项目根目录
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile,"/excel_template/userList.xlsx");
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
            rowIndex ++;
        }

        // 删除第二个sheet
        workbook.removeSheetAt(1);

        // 4、导出文件
        String filename="员工数据.xlsx";
        response.setHeader("content-disposition","attachment;filename=" + new String(filename.getBytes(),"ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        workbook.write(response.getOutputStream());
        workbook.close();
    }

    /**
     * 使用模板导出用户详细数据
     * @param id
     * @param response
     */
    public void downloadUserInfoByTemplate(Long id, HttpServletResponse response) throws IOException, InvalidFormatException {
        // 1、读取模板
        File rootFile = new File(ResourceUtils.getURL("classpath:").getPath());
        File templateFile = new File(rootFile,"/excel_template/userInfo.xlsx");
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
        // 城市    第7行第4列
        sheet.getRow(6).getCell(3).setCellValue(user.getCity());
        // 照片    第2行至第5行，第3列至第4列

        // 4、导出
        String filename="员工("+user.getUserName()+")详细信息.xlsx";
        response.setHeader("content-disposition","attachment;filename=" + new String(filename.getBytes(),"ISO-8859-1"));
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8");
        workbook.write(response.getOutputStream());
        workbook.close();
    }
}
