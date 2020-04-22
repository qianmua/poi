package pres.hjc.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

/**
 * Created by IntelliJ IDEA.
 *
 * @author HJC
 * @version 1.0
 * To change this template use File | Settings | File Templates.
 * @date 2020/4/22
 * @time 19:33
 */
public class Write_1 {
    public void test_03(){

        // 创建工作溥
        Workbook workbook = new HSSFWorkbook();
        // 07版
        Workbook workbook2 = new XSSFWorkbook();
        // 07 升级版
        Workbook workbook3 = new SXSSFWorkbook();

        //创建工作表
        Sheet sheet = workbook.createSheet("测试表");
        //创建行
        Row row1 = sheet.createRow(0);
        //创建单元格
        Cell cell1 = row1.createCell(0);
        Cell cell2 = row1.createCell(1);
        cell2.setCellValue(11111);

        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("时间");
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

    }
}
