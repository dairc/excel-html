package test;

import util.ExcelTransfer;

import java.io.FileNotFoundException;
import java.lang.reflect.InvocationTargetException;

/**
 * 测试类
 * 可以将转换后的字符串粘贴到 html文件中查看效果
 */
public class ExcelToTableTest {
    public static void main(String[] args) throws ClassNotFoundException, NoSuchMethodException, InstantiationException, IllegalAccessException, InvocationTargetException, FileNotFoundException {
        // 创建转换器，参数为Excel文件的路径
        ExcelTransfer toHtml = new ExcelTransfer("E:/excelTest/xxxx.xlsx");
        // 将Excel转换为一个表格标签字符串
        String table = toHtml.toTable();
        System.out.println(table);
    }
}
