package excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Desciption:
 * Author: JasonHan.
 * Creation time: 2017/03/20 09:40.
 * © Copyright 2013-2017, Node Supply Chain Management.
 */
public class ExcelHelper {

    private static final String FIRST_COLUMN = "firstColumn";
    private static final String LAST_COLUMN = "lastColumn";
    private static final String FIRST_ROW = "firstRow";
    private static final String LAST_ROW = "lastRow";

    public static int total;

    public String getExcelPath() throws IOException {
        String excelFilePath = null;
        Properties properties = new Properties();

        InputStream in = this.getClass().getResourceAsStream("../config.properties");
        properties.load(in);
        in.close();

        //遍历配置文件方法一
        Enumeration en = properties.propertyNames();
        while (en.hasMoreElements()) {
            String key = en.nextElement().toString();
            String value = properties.getProperty(key);
            System.out.println(key + ":" + value);
            if (key.equals("fileName")) {
                excelFilePath = value;
            }
        }
        //遍历配置文件方法二
//        Set keys = props.keySet();
//        for (Interator it = keys.iterator(); it.hasNext();){
//            String k = it.next();
//            System.out.println(k + ":" + props.getProperty(k));
//        }
        return excelFilePath;
    }

    public ArrayList<String> loadExcel() throws IOException {
        XSSFWorkbook hssfWorkbook = null;
        ArrayList<String> result = new ArrayList<String>();

        String excelFilePath = getExcelPath();
        System.out.println("excelFilePath:"+excelFilePath);
        if (null!=excelFilePath) {
            System.out.println("excelFilePath:"+excelFilePath);
            InputStream in = this.getClass().getResourceAsStream(excelFilePath);
            hssfWorkbook = new XSSFWorkbook(in);
            in.close();
        }

        result.addAll(this.getSql(hssfWorkbook));
        hssfWorkbook.close();
        return result;
    }

    public ArrayList<String> loadExcel(int startIndex, int endIndex) throws IOException {
        XSSFWorkbook hssfWorkbook = null;
        ArrayList<String> result = new ArrayList<String>();

        String excelFilePath;
        for (int i = startIndex; i <= endIndex; i++) {
            excelFilePath = "../"+i+".xlsx";
            System.out.println("excelFilePath:"+excelFilePath);
            if (null != excelFilePath) {
                InputStream in = this.getClass().getResourceAsStream(excelFilePath);
                hssfWorkbook = new XSSFWorkbook(in);
                in.close();
            }
            result.addAll(this.getSql(hssfWorkbook));
            hssfWorkbook.close();
        }

        return result;
    }

    public ArrayList<String> getSql(XSSFWorkbook hssfWorkbook) {
        ArrayList<String> result = new ArrayList<String>();
        Sheet sheet = hssfWorkbook.getSheetAt(0);
        int rowIndex = 0;
        int cellIndex = 0;
        total += (sheet.getLastRowNum()-1);
        System.out.println("sheet.getLastRowNum():"+sheet.getLastRowNum());
        for (Row row : sheet) {
            cellIndex = 0;
            String step = new String();
            /*for (Cell cell : row) {
//                step.put(cell.getColumnIndex(), getRegionValue(sheet, rowIndex, cellIndex));
                System.out.println(cell.getRowIndex()+" "+cell.getColumnIndex());
                if (cell.getColumnIndex() <= 4) {
                    step.put(cell.getColumnIndex(), cell.getStringCellValue());
                }
//                step.add(cell.getColumnIndex()+":"+getRegionValue(sheet, rowIndex, cellIndex));
                cellIndex++;
            }*/

            int sqlCellnum = 26;
            if (null!=row.getCell(sqlCellnum)) {
//                System.out.println(row.getRowNum()+":"+row.getCell(sqlCellnum).getStringCellValue());
                /*if (row.getRowNum()!=0) {
                    System.out.println(row.getCell(sqlCellnum).getCachedFormulaResultType());
                }*/
                /*switch (row.getCell(sqlCellnum).getCellType()) {
                    case XSSFCell.CELL_TYPE_FORMULA:
                        System.out.println(row.getCell(sqlCellnum).getNumericCellValue());
                        break;
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        System.out.println(row.getCell(sqlCellnum).getNumericCellValue());
                        break;
                    case XSSFCell.CELL_TYPE_STRING:
                        System.out.println(row.getCell(sqlCellnum).getRichStringCellValue());
                        break;
                    default:
                        System.out.println("`````");
                }*/
                /*if (row.getRowNum()!=0) {
                    FormulaEvaluator evaluator = hssfWorkbook.getCreationHelper().createFormulaEvaluator();
                    step = evaluator.evaluate(row.getCell(sqlCellnum)).getStringValue();
                }*/
                FormulaEvaluator evaluator = hssfWorkbook.getCreationHelper().createFormulaEvaluator();
                step = evaluator.evaluate(row.getCell(sqlCellnum)).getStringValue();
//                System.out.printf("%3d"+"   "+"%s\n",row.getRowNum(),evaluator.evaluate(row.getCell(sqlCellnum)).getStringValue());
//                System.out.println(evaluator.evaluate(row.getCell(sqlCellnum)).getStringValue());

                /*if (StringUtils.isNotBlank(row.getCell(sqlCellnum).toString())) {
                    System.out.println(row.getRowNum()+":"+row.getCell(sqlCellnum).getCellFormula());
                }*/
            } else {
                System.out.println(row.getRowNum()+":");
            }

            /*row.getCell(86).getCellFormula();*/

            result.add(step);
            rowIndex++;
        }
        return result;
    }

    public static boolean isMergedRegion(Sheet sheet, int row, int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    return true;
                }
            }
        }
        return false;
    }

    public static Map<String, String> mergedRegionStatus(Sheet sheet, int row, int column) {
        Map<String, String> result = new HashMap<String, String>();
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    result.put("state", "true");
                    result.put(FIRST_ROW, String.valueOf(firstRow));
                    result.put(LAST_ROW, String.valueOf(lastRow));
                    result.put(FIRST_COLUMN, String.valueOf(firstColumn));
                    result.put(LAST_COLUMN, String.valueOf(lastColumn));
                    return result;
                }
            }
        }
        result.put("state", "false");
        return result;
    }

    public static String getRegionValue(Sheet sheet, int row, int column) {
        Cell cellResult = null;
        Map<String, String> source = mergedRegionStatus(sheet, row, column);
        if (source.get("state").equals("true")) {
            Row rowResult = sheet.getRow(Integer.valueOf(source.get(FIRST_ROW)));
            cellResult = rowResult.getCell(Integer.valueOf(source.get(FIRST_COLUMN)));
        } else {
            Row rowResult = sheet.getRow(row);
            if (null == rowResult) {
                System.out.println("rowResult is NULL");
            } else {
                cellResult = rowResult.getCell(column);
            }
            
        }
        if (null != cellResult) {
            return cellResult.getStringCellValue();
        } else {
            return "";
        }

    }

    private boolean isMergedRow(Sheet sheet,int row ,int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row == firstRow && row == lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    return true;
                }
            }
        }
        return false;
    }

    public static ArrayList<Workbook> createWorkBookList(List<Map<String, Object>> list, String[] keys, String columnNames[]) {
        ArrayList<Workbook> workbookList = new ArrayList<Workbook>();
        int listSize = list.size();
        int thresholdValue = 32000;
        int groupSize = list.size()/thresholdValue;
        for (int i = 0; i <= groupSize; i++) {
            List<Map<String, Object>> listNew = new ArrayList<Map<String, Object>>();
            for (int j = 0; j < thresholdValue; j++) {
                if ((j+i*thresholdValue) < listSize) {
                    listNew.add(list.get(j+i*groupSize));
                } else {
                    break;
                }
            }
            Workbook wb = createWorkBook(listNew, keys, columnNames);
            workbookList.add(wb);
        }
        return workbookList;
    }

    private static Workbook createWorkBook(List<Map<String, Object>> list, String []keys, String columnNames[]) {
        // 创建excel工作簿
        Workbook wb = new HSSFWorkbook();
        // 创建第一个sheet（页），并命名
        Sheet sheet = wb.createSheet("物流报价");
        // 手动设置列宽。第一个参数表示要为第几列设；，第二个参数表示列的宽度，n为列高的像素数。
        for(int i=0;i<keys.length;i++){
            sheet.setColumnWidth((short) i, (short) (35.7 * 150));
        }

        // 创建第一行
        Row row = sheet.createRow((short) 0);

        // 创建两种单元格格式
        CellStyle cs = wb.createCellStyle();
        CellStyle cs2 = wb.createCellStyle();

        // 创建两种字体
        Font f = wb.createFont();
        Font f2 = wb.createFont();

        // 创建第一种字体样式（用于列名）
        f.setFontHeightInPoints((short) 10);
        f.setColor(IndexedColors.BLACK.getIndex());
        f.setBoldweight(Font.BOLDWEIGHT_BOLD);

        // 创建第二种字体样式（用于值）
        f2.setFontHeightInPoints((short) 10);
        f2.setColor(IndexedColors.BLACK.getIndex());

//        Font f3=wb.createFont();
//        f3.setFontHeightInPoints((short) 10);
//        f3.setColor(IndexedColors.RED.getIndex());

        // 设置第一种单元格的样式（用于列名）
        cs.setFont(f);
        cs.setBorderLeft(CellStyle.BORDER_THIN);
        cs.setBorderRight(CellStyle.BORDER_THIN);
        cs.setBorderTop(CellStyle.BORDER_THIN);
        cs.setBorderBottom(CellStyle.BORDER_THIN);
        cs.setAlignment(CellStyle.ALIGN_CENTER);

        // 设置第二种单元格的样式（用于值）
        cs2.setFont(f2);
        cs2.setBorderLeft(CellStyle.BORDER_THIN);
        cs2.setBorderRight(CellStyle.BORDER_THIN);
        cs2.setBorderTop(CellStyle.BORDER_THIN);
        cs2.setBorderBottom(CellStyle.BORDER_THIN);
        cs2.setAlignment(CellStyle.ALIGN_CENTER);

        //设置列名
        for(int i=0;i<columnNames.length;i++){
            Cell cell = row.createCell(i);
            cell.setCellValue(columnNames[i]);

            cell.setCellStyle(cs);
        }

        //设置每行每列的值
        for (short i = 1; i < list.size(); i++) {
            // Row 行,Cell 方格 , Row 和 Cell 都是从0开始计数的
            // 创建一行，在页sheet上
            Row row1 = sheet.createRow((short) i);
            // 在row行上创建一个方格
            for(short j=0;j<keys.length;j++){
                Cell cell = row1.createCell(j);
                cell.setCellStyle(cs2);
                cell.setCellValue(list.get(i).get(keys[j]) == null?" ": list.get(i).get(keys[j]).toString());
            }
        }
        return wb;
    }
}
