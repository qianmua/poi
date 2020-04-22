package pres.hjc.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by IntelliJ IDEA.
 *
 * @author HJC
 * @version 1.0
 * To change this template use File | Settings | File Templates.
 * @date 2020/4/22
 * @time 22:55
 */
public class 读取不同类型数据 {
    public static void main(String[] args) {

    }

    static String path = "G:\\study\\测试.xlsx";
    private static void read2() throws IOException {
        //得到流
        FileInputStream inputStream = new FileInputStream(path);
        Workbook workbook = new XSSFWorkbook(inputStream);
        //到到第一张表
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        //遍历出标题
        if (row != null){
            int count = row.getPhysicalNumberOfCells();
            for (int i = 0; i < count; i++) {
                Cell cell = row.getCell(i);
                if (cell != null){
                    int typ = cell.getCellType();
                    String value = cell.getStringCellValue();
                    System.out.print("value = " + value);
                }
            }
            System.out.println();
        }

        //得到表中内容
        int count2 = sheet.getPhysicalNumberOfRows();
        for (int i = 1; i < count2; i++) {
            Row row1 = sheet.getRow(i);
            if (row != null){
                //读取列
                int count = row.getPhysicalNumberOfCells();
                for (int i1 = 0; i1 < count; i1++) {
                    Cell cell = row.getCell(i1);
                    if (cell!= null){
                        int type =  cell.getCellType();
                        String value = "";
                        /*switch (type){
                        }*/
                    }
                }
            }
        }
    }
}
