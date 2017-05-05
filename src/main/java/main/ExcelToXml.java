package main;

import excel.ExcelHelper;
import map.StringHelper;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Desciption:
 * Author: JasonHan.
 * Creation time: 2017/03/20 09:34.
 * © Copyright 2013-2017, Node Supply Chain Management.
 */
public class ExcelToXml {

    public static void main(String[] myArgs) {
//        ExcelHelper excelHelper = new ExcelHelper();
//        try {
////            ArrayList<HashMap> result = excelHelper.loadExcel();
//            ArrayList<HashMap> result = null;
//            List<Map<String, Object>> listmap = new ArrayList<Map<String, Object>>();
//            // 列名
//            String columnNames[] = {
//                    "DATA_SRC_URL", "DATA_SRC_ID", "DATA_SRC_ID_TXT", "ARCH_TM", "PTY_ID",
//                    "PTY_CD", "PTY_SHORT_NM_CN","PTY_FULL_NM_CN","PTY_SHORT_NM_EN","PTY_FULL_NM_EN",
//                    "PTY_CAT_CD","PTY_CAT_NM_CN","PTY_PTY_ID","PTY_PTY_CD", "PTY_PTY_SHORT_NM_CN",
//                    "PTY_PTY_FULL_NM_CN","PTY_PTY_SHORT_NM_EN","PTY_PTY_FULL_NM_EN","PTY_PTY_CAT_CD","PTY_PTY_CAT_NM_CN",
//                    "LOC_ID","LOC_CD","LOC_NM_CN","LOC_NM_EN","MEMO",
//                    "CAT_CD","CAT_NM_CN","GIS_CD","GIS_NM_CN","CTRY_CD",
//                    "CTRY_NM_CN", "ST_CD", "ST_NM", "CTY_CD","CTY_NM_CN",
//                    "DIST_CD", "DIST_NM_CN","TOWN_CD","TOWN_NM_CN", "STR_CD",
//                    "STR_NM_CN", "ADDR_LINE1","ADDR_LINE2","ZIP_CD", "CONTC_NM1",
//                    "CONTC_NM2", "TELEPHONE1", "TELEPHONE2", "MOBILE_PHONE1", "MOBILE_PHONE2",
//                    "FAX1", "FAX2", "EMAIL1", "EMAIL2", "WECHAT",
//                    "QQ" , "DELIV_OTTR", "GIFT_OTTR", "RTRN_NOTE_OTTR", "ORD_FREQ",
//                    "ORD_FREQ_MU_CD" , "ORD_FREQ_MU_NM_CN", "MIN_DELIV_QTY", "MIN_DELIV_QTY_MU_CD", "MIN_DELIV_QTY_MU",
//                    "MIN_GIFT_QTY", "MIN_GIFT_QTY_MU_CD", "MIN_GIFT_QTY_MU_NM_CN", "EXPI_RATIO", "FCLTY_CAT_CD",
//                    "FCLTY_CAT_NM_CN", "VLD_FROM", "VLD_TO", "STAT_CD", "STAT_NM_CN",
//                    "PREF_FLAG", "CRT_TM", "CRT_BY", "CRT_CHL_CD", "UPD_TM",
//                    "UPD_BY", "UPD_CHL_CD", "EDIT_FLAG", "RELS_NBR"};
//            // map中的key
//            String keys[] = {
//                    "DATA_SRC_URL", "DATA_SRC_ID", "DATA_SRC_ID_TXT", "ARCH_TM", "PTY_ID",
//                    "PTY_CD", "PTY_SHORT_NM_CN","PTY_FULL_NM_CN","PTY_SHORT_NM_EN","PTY_FULL_NM_EN",
//                    "PTY_CAT_CD","PTY_CAT_NM_CN","PTY_PTY_ID","PTY_PTY_CD", "PTY_PTY_SHORT_NM_CN",
//                    "PTY_PTY_FULL_NM_CN","PTY_PTY_SHORT_NM_EN","PTY_PTY_FULL_NM_EN","PTY_PTY_CAT_CD","PTY_PTY_CAT_NM_CN",
//                    "LOC_ID","LOC_CD","LOC_NM_CN","LOC_NM_EN","MEMO",
//                    "CAT_CD","CAT_NM_CN","GIS_CD","GIS_NM_CN","CTRY_CD",
//                    "CTRY_NM_CN", "ST_CD", "ST_NM", "CTY_CD","CTY_NM_CN",
//                    "DIST_CD", "DIST_NM_CN","TOWN_CD","TOWN_NM_CN", "STR_CD",
//                    "STR_NM_CN", "ADDR_LINE1","ADDR_LINE2","ZIP_CD", "CONTC_NM1",
//                    "CONTC_NM2", "TELEPHONE1", "TELEPHONE2", "MOBILE_PHONE1", "MOBILE_PHONE2",
//                    "FAX1", "FAX2", "EMAIL1", "EMAIL2", "WECHAT",
//                    "QQ" , "DELIV_OTTR", "GIFT_OTTR", "RTRN_NOTE_OTTR", "ORD_FREQ",
//                    "ORD_FREQ_MU_CD" , "ORD_FREQ_MU_NM_CN", "MIN_DELIV_QTY", "MIN_DELIV_QTY_MU_CD", "MIN_DELIV_QTY_MU",
//                    "MIN_GIFT_QTY", "MIN_GIFT_QTY_MU_CD", "MIN_GIFT_QTY_MU_NM_CN", "EXPI_RATIO", "FCLTY_CAT_CD",
//                    "FCLTY_CAT_NM_CN", "VLD_FROM", "VLD_TO", "STAT_CD", "STAT_NM_CN",
//                    "PREF_FLAG", "CRT_TM", "CRT_BY", "CRT_CHL_CD", "UPD_TM",
//                    "UPD_BY", "UPD_CHL_CD", "EDIT_FLAG", "RELS_NBR"};
//
//            String oldCity = null;
//            for (HashMap stepList : result) {
//                for (int i = 0; i < 7; i++) {
//                    if (null != stepList.get(i)) {
//                        System.out.printf("|%50s", StringHelper.formatStr(stepList.get(i).toString()));
//                    } else {
//                        System.out.printf("|%50s", "");
//                    }
//                }
//
//                if (null != stepList.get(1)) {
//                    Map<String, Object> mapValue = new HashMap<String, Object>();
//                    mapValue.put("DATA_SRC_URL", "157.达能郑州网点_2017.04.12.xlsx");
//                    mapValue.put("DATA_SRC_ID", "");
//                    mapValue.put("DATA_SRC_ID_TXT", "");
//                    mapValue.put("ARCH_TM", "now()");
//                    mapValue.put("PTY_ID", "");
//                    mapValue.put("PTY_CD", "DANONE");
//                    mapValue.put("PTY_SHORT_NM_CN", "达能集团");
//                    mapValue.put("PTY_FULL_NM_CN", "达能集团");
//                    mapValue.put("PTY_SHORT_NM_EN", "DaNeng");
//                    mapValue.put("PTY_FULL_NM_EN", "DaNeng");
//                    mapValue.put("PTY_CAT_CD", "1170.205");
//                    mapValue.put("PTY_CAT_NM_CN", "货主");
//                    mapValue.put("PTY_PTY_ID", "");
//                    mapValue.put("PTY_PTY_CD", "");
//                    mapValue.put("PTY_PTY_SHORT_NM_CN", stepList.get(3));
//                    mapValue.put("PTY_PTY_FULL_NM_CN", stepList.get(3));
//                    mapValue.put("PTY_PTY_SHORT_NM_EN", "");
//                    mapValue.put("PTY_PTY_FULL_NM_EN", "");
//                    mapValue.put("PTY_PTY_CAT_CD", "1170.220.205");
//                    mapValue.put("PTY_PTY_CAT_NM_CN", "经销商");
//                    mapValue.put("LOC_ID", "");
//                    mapValue.put("LOC_CD", "");
//                    mapValue.put("LOC_NM_CN", "");
//                    mapValue.put("LOC_NM_EN", "");
//                    mapValue.put("MEMO", "");
//                    mapValue.put("CAT_CD", "1220.210.300");
//                    mapValue.put("CAT_NM_CN", "货主销售网点收货地址");
//                    mapValue.put("GIS_CD", "");
//                    mapValue.put("GIS_NM_CN", "");
//                    mapValue.put("CTRY_CD", "");
//                    mapValue.put("CTRY_NM_CN", "");
//                    mapValue.put("ST_CD", "");
//                    mapValue.put("ST_NM", "河南");
//                    mapValue.put("CTY_CD", "");
//                    /*if (null == stepList.get(0)) {
//                        oldCity = stepList.get(0).toString();
//                    }*/
//                    mapValue.put("CTY_NM_CN", stepList.get(0));
//                    mapValue.put("DIST_CD", "");
//                    mapValue.put("DIST_NM_CN", "");
//                    mapValue.put("TOWN_CD", "");
//                    mapValue.put("TOWN_NM_CN", "");
//                    mapValue.put("STR_CD", "");
//                    mapValue.put("STR_NM_CN", "");
//                    mapValue.put("ADDR_LINE1", stepList.get(4));
//                    mapValue.put("ADDR_LINE2", "");
//                    mapValue.put("ZIP_CD", "");
//                    mapValue.put("CONTC_NM1", "");
//                    mapValue.put("CONTC_NM2", "");
//                    mapValue.put("TELEPHONE1", "");
//                    mapValue.put("TELEPHONE2", "");
//                    mapValue.put("MOBILE_PHONE1", "");
//                    mapValue.put("MOBILE_PHONE2", "");
//                    mapValue.put("FAX1", "");
//                    mapValue.put("FAX2", "");
//                    mapValue.put("EMAIL1", "");
//                    mapValue.put("EMAIL2", "");
//                    mapValue.put("WECHAT", "");
//                    mapValue.put("QQ", "");
//                    mapValue.put("DELIV_OTTR", "24");
//                    mapValue.put("GIFT_OTTR", "48");
//                    mapValue.put("RTRN_NOTE_OTTR", "72");
//                    mapValue.put("ORD_FREQ", "3");
//                    mapValue.put("ORD_FREQ_MU_CD", "4110.155.215");
//                    mapValue.put("ORD_FREQ_MU_NM_CN", "周");
//                    mapValue.put("MIN_DELIV_QTY", "");
//                    mapValue.put("MIN_DELIV_QTY_MU_CD", "");
//                    mapValue.put("MIN_DELIV_QTY_MU", "");
//                    mapValue.put("MIN_GIFT_QTY", "");
//                    mapValue.put("MIN_GIFT_QTY_MU_CD", "");
//                    mapValue.put("MIN_GIFT_QTY_MU_NM_CN", "");
//                    mapValue.put("EXPI_RATIO", "0.33");
//                    mapValue.put("FCLTY_CAT_CD", "");
//                    mapValue.put("FCLTY_CAT_NM_CN", "");
//                    mapValue.put("VLD_FROM", "");
//                    mapValue.put("VLD_TO", "");
//                    mapValue.put("STAT_CD", "");
//                    mapValue.put("STAT_NM_CN", "");
//                    mapValue.put("PREF_FLAG", "");
//                    mapValue.put("CRT_TM", "now()");
//                    mapValue.put("CRT_BY", "hanzhe");
//                    mapValue.put("CRT_CHL_CD", "");
//                    mapValue.put("UPD_TM", "");
//                    mapValue.put("UPD_BY", "");
//                    mapValue.put("UPD_CHL_CD", "");
//                    mapValue.put("EDIT_FLAG", "1100.200");
//                    mapValue.put("RELS_NBR", "1.0.0");
//
//                    listmap.add(mapValue);
//                }
//
//                System.out.println("------");
//            }
//
//            ArrayList<Workbook> resultList = ExcelHelper.createWorkBookList(listmap, keys, columnNames);
//
//            try {
//                if (resultList.size() == 1) {
//                    FileOutputStream fout = new FileOutputStream("D:\\HEIHEIxiawu"+".xls");
//                    resultList.get(0).write(fout);
//                    fout.close();
//                } else {
//                    int mark = 1;
//                    for (Workbook step : resultList) {
//                        FileOutputStream fout = new FileOutputStream("D:\\HEIHEIxiawu"+"-"+mark+".xls");
//                        step.write(fout);
//                        fout.close();
//                        mark++;
//                    }
//                }
//            } catch (IOException e) {
//                e.printStackTrace();
//            }
//
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
    }

    public static String changeCharset(String str, String newCharset){
        if (str != null) {
            //用默认字符编码解码字符串。
            byte[] bs = str.getBytes();
            //用新的字符编码生成字符串
            try {
                return new String(bs, newCharset);
            } catch (UnsupportedEncodingException e) {
                e.printStackTrace();
            }
        }
        return null;
    }
}
