package pres.hjc.poi;

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
 * @time 22:44
 */
public class 读取数据 {
    public static void main(String[] args) throws IOException {
        read();
    }
    static  String path = "G:\\study\\";
    private static void read() throws IOException {
        //得到文件流
        FileInputStream inputStream = new FileInputStream(path + "测试1.xlsx");
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);

        System.out.println(cell.getNumericCellValue());
    }
}
