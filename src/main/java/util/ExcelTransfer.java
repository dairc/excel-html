package util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
/**
 * 将Excel 文件转换为 html表格 table
 * toTable 方法为转换方法， 使用的都是接口，所以具体使用什么样的实现并不影响
 */
public class ExcelTransfer {


    /**
     * 两种不同的工作簿，对应了Excel2003和Excel2007
     * 它们都实现了Workbook
     */
    private Workbook workbook;
    // Excel文件路径
    private String filePath;

    // 构造方法
    public ExcelTransfer() {}
    // 构造方法，有参，会初始化workbook
    public ExcelTransfer(String filePath) throws ClassNotFoundException, NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException, FileNotFoundException {
        this.filePath = filePath;
        initWorkbookType();
    }

    // getter setter
    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public String getFilePath() {
        return filePath;
    }

    public void setFilePath(String filePath) {
        this.filePath = filePath;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    /**
     * 将excel转换成表格 table
     * @return table 的字符串
     */
    public String toTable() {
        StringBuilder table = new StringBuilder();
        table.append("<table cellspacing='0' border='1'>");
        Sheet sheet = workbook.getSheetAt(0);
        if (sheet != null) {
            Iterator<Row> rowIterator = sheet.rowIterator();

            while (rowIterator.hasNext()) {
                table.append("<tr");
                Row row = rowIterator.next();

                table.append(">");
                Iterator<Cell> cellIterator = row.cellIterator();
                // 遇到合并列时进行计数
                int colCount = 0;
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    // 判断是否为合并单元格
                    Map<String, Integer> map = getMergedRegionInfo(sheet, cell);

                    // 如果有合并列
                    // 如果这一列有colspan属性（是一个合并列），并且是第一次被获取到，
                    // 那么进行计数器的增加，防止第二次获取到同一个合并单元格时，进行单元格的重复添加
                    if (map.get("colspan") != null) {
                        if (colCount++ == 0) {
                            table.append("<td ");
                            table.append("colspan=" + map.get("colspan"));
                            table.append(">");
                            table.append(cell.getStringCellValue());
                            table.append("</td>");
                        }
                    } else if (map.get("rowspan") != null) {// 如果有合并行
                        // 如果这一列的当前行就是这个合并单元格的第一行，那么我们创建这个单元格，并合并对应的行数
                        int rowIndex = cell.getRowIndex();
                        if (rowIndex++ == map.get("firstR")) {
                            table.append("<td ");
                            table.append("rowspan=" + map.get("rowspan"));
                            table.append(">");
                            table.append(cell.getStringCellValue());
                            table.append("</td>");
                        }
                    } else {
                        // 如果没有合并列
                        table.append("<td>");
                        table.append(cell.getStringCellValue());
                        table.append("</td>");
                    }
                }
                table.append("</tr>");
            }
            table.append("</table>");

            return table.toString();
        }

        return null;
    }

    /**
     * 初始化workbook
     */
    private void initWorkbookType() throws ClassNotFoundException, NoSuchMethodException,
            IllegalAccessException, InvocationTargetException, InstantiationException, FileNotFoundException {
        // 根据文件类型，决定使用的工作簿的类型
        String type = "xls".equals(getSuffix().trim()) ?
                "org.apache.poi.hssf.usermodel.HSSFWorkbook" : "org.apache.poi.xssf.usermodel.XSSFWorkbook";

        // 文件输入流读取文件
        InputStream in = new FileInputStream(filePath);
        // 反射创建workbook
        Class workbookClass = Class.forName(type);
        this.workbook = (Workbook) workbookClass.getConstructor(InputStream.class).newInstance(in);
    }

    /**
     * 获取文件后缀名
     * @return
     */
    private String getSuffix() {
        int dotIndex = filePath.lastIndexOf(".");
        return filePath.substring(dotIndex + 1);
    }

    /**
     * 读取合并的单元格信息
     * @param sheet
     * @param cell
     * @return
     */
    private Map<String, Integer> getMergedRegionInfo(Sheet sheet, Cell cell) {
        int firstC;
        int lastC;
        int firstR;
        int lastR;
        Map<String, Integer> map = new HashMap<>();
        for (int cellNum = 0; cellNum < sheet.getNumMergedRegions(); cellNum++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(cellNum);
            firstC = mergedRegion.getFirstColumn();
            lastC = mergedRegion.getLastColumn();
            firstR = mergedRegion.getFirstRow();
            lastR = mergedRegion.getLastRow();

            // 如果这个单元格在这个 合并区域的范围内，那么它就是这个单元格
            if (cell.getRowIndex() >= firstR && cell.getRowIndex() <= lastR) {
                if (cell.getColumnIndex() >= firstC && cell.getColumnIndex() <= lastC) {
                    // 合并的合数和列数，以及合并单元格的第一行索引
                    int rowspan = lastR-firstR;
                    int colspan = lastC-firstC;
                    if (rowspan > 0)
                        map.put("rowspan", rowspan+1);
                    map.put("firstR", firstR);
                    if (colspan > 0)
                        map.put("colspan", colspan+1);
                }
            }
        }
        return map;
    }
}