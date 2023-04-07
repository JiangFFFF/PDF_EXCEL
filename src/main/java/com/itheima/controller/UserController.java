package com.itheima.controller;

import com.itheima.pojo.User;
import com.itheima.service.UserService;
import jxl.write.WriteException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
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

}
