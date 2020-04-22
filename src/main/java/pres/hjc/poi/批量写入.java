package pres.hjc.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by IntelliJ IDEA.
 *
 * @author HJC
 * @version 1.0
 * To change this template use File | Settings | File Templates.
 * @date 2020/4/22
 * @time 22:24
 */
public class 批量写入 {
    public static void main(String[] args) throws IOException {
        allWrite();
    }

    static String path = "G:\\study\\";

    public static void allWrite() throws IOException {
        long s = System.currentTimeMillis();
        // 07 版 xlsx
        Workbook workbook = new XSSFWorkbook();
        // 03 版 xls
        Workbook workbook2 = new HSSFWorkbook();
        // 优化版 需要关闭临时文件
        Workbook workbook3 = new SXSSFWorkbook();
        Sheet sheet = workbook.createSheet("测试");
        for (int i = 0; i < 65537; i++) {
            Row row = sheet.createRow(i);
            for (int i1 = 0; i1 < 10; i1++) {
                Cell cell = row.createCell(i1);
                cell.setCellValue(i1);
            }
        }
        System.out.println("写入结束。");
        FileOutputStream outputStream = new FileOutputStream(path + "测试1.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        //清理掉临时文件
        //((SXSSFWorkbook) workbook3).dispose();
        long e = System.currentTimeMillis();
        System.out.println((e - s )+ " ms");
    }

}
