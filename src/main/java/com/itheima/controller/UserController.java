package com.itheima.controller;

import com.itheima.pojo.User;
import com.itheima.service.UserService;
import jxl.write.WriteException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
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
            userService.uploadExcel(file);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (ParseException e) {
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
            userService.downloadUserInfoByTemplate(id,response);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

}
