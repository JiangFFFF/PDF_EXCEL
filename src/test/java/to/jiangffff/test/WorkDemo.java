package to.jiangffff.test;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

/**
 * @author JiangHuifeng
 * @create 2023-04-15-14:54
 */
public class WorkDemo {
    public static void main(String[] args) throws IOException {
        XWPFDocument document = new XWPFDocument(new FileInputStream("/Users/jianghuifeng/Desktop/基础教程/java报表/day3/资料/test.docx"));
        // 读取正文
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                System.out.println(run.getText(0));
            }
        }

        // 读取表格
        XWPFTable table = document.getTables().get(0);
        List<XWPFTableRow> rows = table.getRows();
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> tableCells = row.getTableCells();
            for (XWPFTableCell tableCell : tableCells) {
                List<XWPFParagraph> xwpfParagraphs = tableCell.getParagraphs();
                for (XWPFParagraph xwpfParagraph : xwpfParagraphs) {
                    System.out.println(xwpfParagraph.getText());
                }
            }
        }


    }
}
