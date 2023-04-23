package to.jiangffff.test;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.data.general.DefaultPieDataset;

import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 *  饼状图
 * @author JiangHuifeng
 * @create 2023-04-16-18:18
 */
public class JFreeChartsDemo1 {
    public static void main(String[] args) throws IOException {
        // 统计每个部门的人员
        DefaultPieDataset dataset = new DefaultPieDataset();
        dataset.setValue("技术部",180);
        dataset.setValue("销售部",20);
        dataset.setValue("人事部",10);

        StandardChartTheme chartTheme = new StandardChartTheme("CN");
        // 设置大标题字体
        chartTheme.setExtraLargeFont(new Font("宋体",Font.BOLD,20));
        // 设置图例字体
        chartTheme.setRegularFont(new Font("宋体",Font.BOLD,12));
        // 设置内容字体
        chartTheme.setLargeFont(new Font("宋体",Font.BOLD,12));
        ChartFactory.setChartTheme(chartTheme);

        /**
         * String title,                    大标题
         * PieDataset dataset,              数据集
         * PieDataset previousDataset,
         * int percentDiffForMaxScale,
         * boolean greenForIncrease,
         * boolean legend,                  是否显示图例
         * boolean tooltips,                是否显示提示
         * boolean urls,                    是否跳转
         * boolean subTitle,
         * boolean showDifference
         */
        JFreeChart chart = ChartFactory.createPieChart3D("部门人员统计", dataset, true, false, false);
        ChartUtils.writeChartAsPNG(new FileOutputStream("/Users/jianghuifeng/Downloads/chart1.png"),chart,400,300);
    }
}
