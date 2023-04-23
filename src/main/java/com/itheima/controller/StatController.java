package com.itheima.controller;

import com.itheima.service.StatService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.List;
import java.util.Map;

/**
 * @author JiangHuifeng
 * @create 2023-04-17-22:33
 */
@RestController
@RequestMapping("/stat")
public class StatController {

    @Autowired
    private StatService statService;

    /**
     * 统计各部门人数
     * @return
     */
    @GetMapping("/columnCharts")
    public List<Map> columnCharts(){
        return statService.columnCharts();
    }

    /**
     * 月份入职人数统计
     * @return
     */
    @GetMapping("/lineCharts")
    public List<Map> lineCharts(){
        return statService.lineCharts();
    }

    /**
     * 员工地方来源统计
     * @return
     */
    @GetMapping("/pieCharts")
    public List<Map<String, Object>> pieCharts(){
        return statService.pieCharts();
    }

    /**
     * echarts饼图
     * @return
     */
    @GetMapping("/pieECharts")
    public Map<String, Object> pieECharts(){
        return statService.pieECharts();
    }

}
