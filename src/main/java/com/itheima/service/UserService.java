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
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;


import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
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
}
