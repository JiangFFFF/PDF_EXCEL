package to.jiangffff.test;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 *  折线图
 * @author JiangHuifeng
 * @create 2023-04-16-18:18
 */
public class JFreeChartsDemo2 {
    public static void main(String[] args) throws IOException {
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
        chartTheme.setExtraLargeFont(new Font("宋体",Font.BOLD,20));
        // 设置图例字体
        chartTheme.setRegularFont(new Font("宋体",Font.BOLD,12));
        // 设置内容字体
        chartTheme.setLargeFont(new Font("宋体",Font.BOLD,12));
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
        JFreeChart chart = ChartFactory.createLineChart("公司人数","各部门","入职人数", dataset);
        ChartUtils.writeChartAsPNG(new FileOutputStream("/Users/jianghuifeng/Downloads/chart2.png"),chart,400,300);
    }
}
