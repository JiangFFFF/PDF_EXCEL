package com.itheima.service;

import com.itheima.mapper.UserMapper;
import com.itheima.pojo.User;
import org.springframework.stereotype.Service;
import tk.mybatis.mapper.entity.Example;

import javax.annotation.Resource;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author JiangHuifeng
 * @create 2023-04-17-22:34
 */
@Service
public class StatService {

    @Resource
    private UserMapper userMapper;

    /**
     * 统计各部门人数
     * @return
     */
    public List<Map> columnCharts() {
        return userMapper.columnCharts();
    }

    /**
     * 月份入职人数统计
     * @return
     */
    public List<Map> lineCharts() {
        return userMapper.lineCharts();
    }

    /**
     * 员工地方来源统计
     * @return
     */
    public List<Map<String, Object>> pieCharts() {
        List<Map<String,Object>> resultMapList = new ArrayList<>();
        //最终想要的数据格式：[{id:"河北省","drilldown":"河北省","name":"河北省","y":10,"data":[{"name":"石家庄","y":4},{"name":"唐山","y":3},{"name":"保定","y":3}]}]
        List<User> userList = userMapper.selectAll();
//        Map<String, Long> collect = userList.stream().collect(Collectors.groupingBy(User::getProvince, Collectors.counting()));
        Map<String, List<User>> provinceMap = userList.stream().collect(Collectors.groupingBy(User::getProvince));
        for (Map.Entry<String, List<User>> province : provinceMap.entrySet()) {
            Map<String,Object> resultMap = new HashMap<>();
            resultMap.put("name",province.getKey());
            resultMap.put("drilldown",province.getKey());
            resultMap.put("id",province.getKey());
            List<User> user = province.getValue();
            resultMap.put("y",user.size());
            List<Map<String,Object>> cityMapList = new ArrayList<>();
            Map<String, Long> citySizeMap = user.stream().collect(Collectors.groupingBy(User::getCity, Collectors.counting()));
            for (Map.Entry<String, Long> cityMap : citySizeMap.entrySet()) {
                Map<String,Object> tempMap = new HashMap<>();
                tempMap.put("name",cityMap.getKey());
                tempMap.put("y",cityMap.getValue());
                cityMapList.add(tempMap);
            }
            resultMap.put("data",cityMapList);
            resultMapList.add(resultMap);
        }
        return resultMapList;
    }

    public Map<String, Object> pieECharts() {
        Map resultMap = new HashMap();
        List<Map> provinceMapList = new ArrayList<>();
        List<Map> cityMapList = new ArrayList<>();
        //        最终想要的数据格式：{province:[{name:,value:}],city:[]}
        Example example = new Example(User.class);
        example.setOrderByClause("province,city");
        List<User> userList = userMapper.selectByExample(example);
        LinkedHashMap<String, List<User>> provinceMap = userList.stream().collect(Collectors.groupingBy(User::getProvince, LinkedHashMap::new, Collectors.toList()));
        for (String provinceName : provinceMap.keySet()) {
            Map map = new HashMap();
            map.put("name",provinceName);
            map.put("value",provinceMap.get(provinceName).size());
            provinceMapList.add(map);
        }

        //        注意分组时得排序，不然数据会乱
        Map<String, List<User>> cityMap = userList.stream().collect(Collectors.groupingBy(User::getCity, LinkedHashMap::new,Collectors.toList()));
        for (String cityName : cityMap.keySet()) {
            Map map = new HashMap();
            map.put("name",cityName);
            map.put("value",cityMap.get(cityName).size());
            cityMapList.add(map);
        }

        resultMap.put("province",provinceMapList);
        resultMap.put("city",cityMapList);
        return resultMap;
    }
}
