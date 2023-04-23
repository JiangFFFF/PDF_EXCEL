package com.itheima.controller;

import com.itheima.pojo.User;
import com.itheima.service.UserService;
import jxl.write.WriteException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.data.category.DefaultCategoryDataset;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.Font;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.List;

@RestController
@RequestMapping("/user")
public class UserController {

    @Autowired
    private UserService userService;

    @GetMapping("/findPage")
    public List<User>  findPage(
            @RequestParam(value = "page",defaultValue = "1") Integer page,
            @RequestParam(value = "rows",defaultValue = "10") Integer pageSize){
        return userService.findPage(page,pageSize);
    }

    @GetMapping(value = "/downLoadXlsByJxl",name = "使用jxl导出excel")
    public void downLoadXlsByJxl(HttpServletResponse response){
        try {
            userService.downLoadXlsByJxl(response);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    @PostMapping(value = "/uploadExcel",name = "上传用户数据")
    public void uploadExcel(MultipartFile file){
        try {
//            userService.uploadExcel(file);
            userService.uploadExcelWithEasyPOI(file);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ParseException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @GetMapping(value = "downLoadXlsxByPoi",name = "使用poi导出数据")
    public void downLoadXlsxByPoi(HttpServletResponse response){
        try {
            // 不带样式
//            userService.downLoadXlsxByPoi(response);
            // 带样式
//            userService.downLoadXlsxByPoiWithCellStyle(response);
            // 模板导出
            userService.downLoadXlsxByPoiWithTemplate(response);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    @GetMapping(value = "/download",name = "使用模板导出用户详细数据")
    public void downloadUserInfoByTemplate(Long id,HttpServletResponse response){
        try {
//            userService.downloadUserInfoByTemplate(id,response);
//            userService.downloadUserInfoByTemplate2(id,response);
//            userService.downloadUserInfoByEasyPOI(id,response);
            userService.downloadUserInfoByPDF(id,response);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @GetMapping(value = "/downLoadMillion",name = "导出百万数据")
    public void downLoadMillion(HttpServletResponse response){
        try {
            userService.downLoadMillion(response);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @GetMapping(value = "/downLoadCSV",name = "用csv文件导出百万数据")
    public void downLoadCSV(HttpServletResponse response){
        try {
//            userService.downLoadCSV(response);
            userService.downLoadCSVWithEasyPOI(response);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @GetMapping(value = "/{id}",name = "根据id查询用户数据")
    public User findById(@PathVariable("id") Long id){
        return userService.findById(id);
    }

    @GetMapping(value = "/downloadContract",name = "下载用户合同文档")
    public void downloadContract(HttpServletResponse response,@RequestParam Long id){
        try {
//            userService.downloadContract(id,response);
            userService.downloadContractByEasyPOI(id,response);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @GetMapping(value = "/downLoadPDF",name = "导出用户数据到pdf中")
    public void downLoadPDF(HttpServletResponse response){
        try {
//            userService.downLoadPDF(response);
            userService.downLoadPDF2(response);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @GetMapping("/jfreeChart")
    public void jfreeChart(HttpServletResponse response) throws IOException {
        // 统计每年各部门入职的人数
        DefaultCategoryDataset dataset = new DefaultCategoryDataset();
        dataset.setValue(200,"技术部","2011");
        dataset.setValue(250,"技术部","2012");
        dataset.setValue(260,"技术部","2013");
        dataset.setValue(280,"技术部","2014");
        dataset.setValue(275,"技术部","2015");

        dataset.setValue(350,"软件部","2011");
        dataset.setValue(340,"软件部","2012");
        dataset.setValue(320,"软件部","2013");
        dataset.setValue(300,"软件部","2014");
        dataset.setValue(275,"软件部","2015");

        dataset.setValue(50,"销售部","2011");
        dataset.setValue(100,"销售部","2012");
        dataset.setValue(200,"销售部","2013");
        dataset.setValue(1000,"销售部","2014");
        dataset.setValue(800,"销售部","2015");

        dataset.setValue(0,"产品部","2011");
        dataset.setValue(0,"产品部","2012");
        dataset.setValue(100,"产品部","2013");
        dataset.setValue(300,"产品部","2014");
        dataset.setValue(600,"产品部","2015");

        StandardChartTheme chartTheme = new StandardChartTheme("CN");
        // 设置大标题字体
        chartTheme.setExtraLargeFont(new java.awt.Font("宋体", java.awt.Font.BOLD,20));
        // 设置图例字体
        chartTheme.setRegularFont(new java.awt.Font("宋体", java.awt.Font.BOLD,12));
        // 设置内容字体
        chartTheme.setLargeFont(new java.awt.Font("宋体", Font.BOLD,12));
        ChartFactory.setChartTheme(chartTheme);

        /**
         * String title,                大标题
         * String categoryAxisLabel,    X轴说明
         * String valueAxisLabel,       Y轴说明
         * CategoryDataset dataset,     数据集
         * PlotOrientation orientation,
         * boolean legend,
         * boolean tooltips,
         * boolean urls
         */
        JFreeChart chart = ChartFactory.createBarChart("公司人数","各部门","入职人数", dataset);
        ChartUtils.writeChartAsJPEG(response.getOutputStream(),chart,400,300);
    }



}
