package main;

import excel.ExcelHelper;
import map.StringHelper;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import sql.MySQL;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;

/**
 * Desciption:
 * Author: JasonHan.
 * Creation time: 2017/04/13 11:26:00.
 * © Copyright 2013-2017, Node Supply Chain Management.
 */
public class ExecuteSQL {
    public static void main(String[] myArgs) {
        ExcelHelper excelHelper = new ExcelHelper();
        try {
//            for (HashMap stepList : result) {
//                for (int i = 0; i < 7; i++) {
//                    if (null != stepList.get(i)) {
//                        System.out.printf("|%50s", StringHelper.formatStr(stepList.get(i).toString()));
//                    } else {
//                        System.out.printf("|%50s", "");
//                    }
//                }
//            }

            ArrayList<String> result = null;
            for (int i = 1; i <= 35; i++) {
                /*File writename = new File("D:\\sql\\logistics-quotation-"+i+".sql"); // 相对路径
                writename.createNewFile(); // 创建新文件
                BufferedWriter out = new BufferedWriter(new FileWriter(writename));
                result = excelHelper.loadExcel(i, i);
                for (String step : result) {
                    out.write(step+"\r\n"); // win换行
                }
                System.out.println(ExcelHelper.total);
                out.flush(); // 把缓存区内容压入文件
                out.close(); // 最后记得关闭文件*/
                result = excelHelper.loadExcel(i, i);
                for (String step : result) {
                    String[] stepArr = step.split(",");
                    if (stepArr.length > 68) {
                        if (!needChange(stepArr[77])) {
                            if (needChange(stepArr[69])) {
                                stepArr[69] = "'0'";
                            }
                            if (needChange(stepArr[70])) {
                                stepArr[70] = "'0'";
                            }
                            if (needChange(stepArr[73])) {
                                stepArr[73] = "'0'";
                            }
                            if (needChange(stepArr[74])) {
                                stepArr[74] = "'0'";
                            }
                            System.out.printf("minWgt:%9s  maxWgt:%9s  minVol:%9s  maxVol:%9s  amt:%9s|%11s|%s\n",
                                    stepArr[69], stepArr[70],
                                    stepArr[73], stepArr[74],
                                    stepArr[77],stepArr[78],stepArr[79]
                            );
                            /*String newStep = "";
                            for (String newArrStep : stepArr) {
                                newStep = newArrStep + ",";
                            }*/
                            System.out.println(i+" 结果:"+MySQL.insertQuotation(StringUtils.join(stepArr,",")));
                        }
                    } else {
                        System.out.println("残缺："+step);
                    }
//                    int sqlResult = MySQL.insertQuotation(step);
//                    if (0 == sqlResult) {
//                        System.out.println(step);
//                    }
//                    System.out.println(sqlResult);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean needChange(String source) {
        return StringUtils.isBlank(source.replace("'",""));
    }

    /*private static void changeValue(String source) {
        if (needChange(source)) {

        }
    }*/
}
