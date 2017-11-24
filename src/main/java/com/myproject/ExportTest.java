package com.myproject;

import com.myproject.common.C_MONGODB;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by admin on 2017/11/24.
 */
public class ExportTest {
    private class JsonKey{
        private String key;
        private int cellnum;

        public String getKey(){
            return key;
        }

        public int getCellNum(){
            return cellnum;
        }

        public void setKey(String key){
            this.key = key;
        }

        public void setCellNum(int num){
            this.cellnum = num;
        }
    };

    private static HSSFWorkbook readFile(String filename) throws IOException {
        FileInputStream fis = new FileInputStream(filename);
        try {
            return new HSSFWorkbook(fis);
        } finally {
            fis.close();
        }
    }
    public static void main(String[] args) {
        JsonKey JKey[]=null;  //定义表头
        int flag = 0;         //定义表头标志（一般是首行）
        String oid = null ;   //存储表头的MongoDB的“_id”


        if (args.length < 1) {
            System.err.println("At least one argument expected");
            return;
        }

        String fileName = args[0];
        try {
            if (args.length < 2) {
                HSSFWorkbook wb = ExportTest.readFile(fileName);

                System.out.println("Data dump:\n");

                for (int k = 0; k < wb.getNumberOfSheets(); k++) {
                    HSSFSheet sheet = wb.getSheetAt(k);
                    int rows = sheet.getPhysicalNumberOfRows();
                    System.out.println("Sheet " + k + " \"" + wb.getSheetName(k) + "\" has " + rows + " row(s).");
                    StringBuffer tabledef = new StringBuffer();
                    tabledef.append("{\"TableDef\":\"" + wb.getSheetName(k) + "\",");
                    for (int r = 0; r < rows; r++) {
                        HSSFRow row = sheet.getRow(r);
                        //过滤空行
                        if (row == null) {
                            rows++;
                            continue;
                        }else{
                            if (flag==0){
                                int cells = row.getPhysicalNumberOfCells();
                                StringBuffer tableheader = new StringBuffer();
                                tableheader.append("\"Columns\":[");
                                JKey = new JsonKey[cells];
                                int cellnum = 0;
                                //此部门用于取表头
                                for (int c = 0; c < cells; c++) {
                                    HSSFCell cell = row.getCell(c);
                                    //过滤空列
                                    if (cell == null){
                                        cells++;
                                        continue;
                                    }
                                    JKey[cellnum] = new ExportTest().new JsonKey();

                                    String value = null;

                                    value = cell.getStringCellValue();
                                    JKey[cellnum].setKey(value);
                                    JKey[cellnum].setCellNum(c);
                                    tableheader.append("{\"FiledName\":\"" + value + "\"}");
                                    tableheader.append(",");
                                    cellnum ++;
                                }
                                tableheader.deleteCharAt(tableheader.length()-1);
                                tableheader.append("]}");
                                flag = 1;
                                tabledef.append(tableheader);
                                System.out.println("TableDef " + tabledef.toString());
                                //保存到MongoDB
                                oid = C_MONGODB.saveObjectByJson(tabledef.toString(),wb.getSheetName(k));
                                continue;
                            }
                        }

                        int cells = JKey.length;
                        StringBuffer datastr = new StringBuffer();
                        datastr.append("{\"oid\":\"" + oid + "\",");  //标记本次导入数据的表头定义的_id
                        for (int c = 0; c < cells; c++) {
                            HSSFCell cell = row.getCell(JKey[c].getCellNum());

                            String value = null;

                            switch (cell.getCellType()) {
                                case HSSFCell.CELL_TYPE_FORMULA:
                                    value = "\"" + JKey[c].getKey() + "\":\"" + cell.getCellFormula() + "\"";
                                    break;
                                case HSSFCell.CELL_TYPE_NUMERIC:
                                    value = "\"" + JKey[c].getKey() + "\":\"" + cell.getNumericCellValue() + "\"";
                                    break;
                                case HSSFCell.CELL_TYPE_STRING:
                                    value = "\"" + JKey[c].getKey() + "\":\"" + cell.getStringCellValue() + "\"";
                                    break;
                                default:
                            }
                            if (value == null){
                                value = "\"" + JKey[c].getKey() + "\":\"\"";;
                            }
                            datastr.append(value);
                            datastr.append(",");
                        }
                        datastr.deleteCharAt(datastr.length()-1);
                        datastr.append("}");
                        C_MONGODB.saveObjectByJson(datastr.toString(),wb.getSheetName(k));
                        System.out.println(datastr.toString());
                    }
                }
                wb.close();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}