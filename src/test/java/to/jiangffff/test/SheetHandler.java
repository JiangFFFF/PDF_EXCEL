package to.jiangffff.test;

import com.itheima.pojo.User;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * @author JiangHuifeng
 * @create 2023-04-13-21:51
 */
public class SheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

    private User user = null;

    /**
     * 每一行的开始
     *
     * @param rowIndex
     */
    @Override
    public void startRow(int rowIndex) {
        if(rowIndex == 0){
            user = null;
        }else {
            user = new User();
        }
    }

    /**
     * 每一行的结束
     *
     * @param rowIndex
     */
    @Override
    public void endRow(int rowIndex) {
        if(rowIndex != 0){
            System.out.println(user);
        }

    }

    /**
     * 处理每一行的所有单元格
     *
     * @param cellName
     * @param cellValue
     * @param xssfComment
     */
    @Override
    public void cell(String cellName, String cellValue, XSSFComment xssfComment) {
        if(user != null){
            // 获取每个单元格名称的首字母 A B C ...
            String letter = cellName.substring(0, 1);
            switch (letter){
                case "A":
                    user.setId(Long.parseLong(cellValue));
                    break;
                case "B":
                    user.setUserName(cellValue);
                    break;
                default:

            }
        }
    }
}
