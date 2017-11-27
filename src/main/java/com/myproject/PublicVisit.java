package com.myproject;

import com.alibaba.fastjson.JSONArray;
import com.mongodb.*;
import org.apache.poi.hssf.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class PublicVisit {
    private static Mongo mg = null;
    private static DB db;
    private static DBCollection visit;

    public static void init(String ip, int port, String database, String collection) {
        try {
            mg = new Mongo(ip, port);
    /*} catch (UnknownHostException e) {
      e.printStackTrace();*/
        } catch (MongoException e) {
            e.printStackTrace();
        }
        db = mg.getDB(database);
        visit = db.getCollection(collection);
    }

    public static void destory() {
        if (mg != null) {
            mg.close();
            mg = null;
            db = null;
            visit = null;
            System.gc();
        }
    }

    /**
     * 添加查询条件：包含“待查询的公众号”的文件
     * @param filePath
     * @return
     * @throws IOException
     */
    public static List<String> readIdsFromFile(String filePath) throws IOException {
        File file = new File(filePath);
        if (file.exists()) {
            FileInputStream fis = new FileInputStream(file);
            BufferedReader reader = new BufferedReader(new InputStreamReader(fis));
            String line = null;
            List ids = new ArrayList();
            while ((line = reader.readLine()) != null) {
                ids.add(line);
            }
            return ids;
        } else {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Sheet1");
            FileOutputStream fs = new FileOutputStream(filePath);
            workbook.write(fs);
            fs.close();
            return null;
        }
    }

    /**
     * 根据查询条件，从Mongo中获取数据
     * @param idsPath 可将“表头信息”、精确查询的订单号等存入
     *        注：暂时还没有用到！
     * @param startTime
     * @param endTime
     * @return
     * @throws IOException
     */
    /*public static JSONArray getAll(String idsPath, String startTime, String endTime) throws IOException {
        List<String> ids = readIdsFromFile(idsPath);
        long start = Long.valueOf(startTime).longValue();
        long end = Long.valueOf(endTime).longValue();
        JSONArray array = new JSONArray();
        for (String id : ids) {
            BasicDBObject query = new BasicDBObject();
            query.put("_id", id);
            DBCursor cursor = visit.find(query);
            List dbs = null;
            if (cursor != null) {
                dbs = cursor.toArray();
                //dbs.get(0)，获取表头
                JSONObject object = JSONObject.parseObject(((DBObject)dbs.get(0)).toString());
                JSONObject item = new JSONObject();
                VisitDetail visitDetail = new VisitDetail();
                visitDetail.setId(object.getString("_id"));
                JSONArray details = object.getJSONArray("visitDetails");
                ArrayList items = new ArrayList();
                for (Iterator i$ = details.iterator(); i$.hasNext(); ) {
                    Object json = i$.next();
                    LinkedHashMap oneMap = new LinkedHashMap();
                    JSONObject one = (JSONObject)json;
                    long date = Long.valueOf(one.get("date").toString()).longValue();
                    if ((date >= start) && (date <= end)) {
                        oneMap.put("date", Long.valueOf(date));
                        oneMap.put("visitNum", one.get("visitNum"));
                        items.add(oneMap);
                    }
                }
                visitDetail.setVisits(items);
                array.add(visitDetail);
            }
        }
        return array;
    }*/

    /**
     * 根据查询条件，从Mongo中获取数据
     * @return
     * @throws IOException
     */
    public static JSONArray getAll() throws IOException {
        JSONArray array = new JSONArray();
        BasicDBObject query = new BasicDBObject();
        query.put("oid", "5a17d9115e16e636d2062c7b");
        DBCursor cursor = visit.find(query);
        if (cursor != null) {
            array.add(cursor.toArray());
        }
        return array;
    }

    public static void exportAsExcel(JSONArray array, String path) {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("访问记录");
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment((short)2);
        HSSFRow row0 = sheet.createRow(0);
        HSSFCell cell = row0.createCell(0);
        row0.createCell(0).setCellValue("订单号");
        row0.createCell(1).setCellValue("支付时间");
        row0.createCell(2).setCellValue("学员账号");
        row0.createCell(3).setCellValue("SKU");
        row0.createCell(4).setCellValue("班型");
        row0.createCell(5).setCellValue("商品价格");
        row0.createCell(6).setCellValue("实付金额");
        row0.createCell(7).setCellValue("原班级");
        row0.createCell(8).setCellValue("现班级");
        row0.createCell(9).setCellValue("班级异动");
        row0.createCell(10).setCellValue("异常编码");
        cell.setCellStyle(style);
        List<DBObject> dbObjectList = (List<DBObject>)array.get(0);
        for (int i = 0; i < dbObjectList.size(); i++) {
            HSSFRow row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(dbObjectList.get(i).get("订单号").toString());
            row.createCell(1).setCellValue(dbObjectList.get(i).get("支付时间").toString());
            row.createCell(2).setCellValue(dbObjectList.get(i).get("学员账号").toString());
            row.createCell(3).setCellValue(dbObjectList.get(i).get("SKU").toString());
            row.createCell(4).setCellValue(dbObjectList.get(i).get("班型").toString());
            row.createCell(5).setCellValue(dbObjectList.get(i).get("商品价格").toString());
            row.createCell(6).setCellValue(dbObjectList.get(i).get("实付金额").toString());
            row.createCell(7).setCellValue(dbObjectList.get(i).get("原班级").toString());
            row.createCell(8).setCellValue(dbObjectList.get(i).get("现班级").toString());
            row.createCell(9).setCellValue(dbObjectList.get(i).get("班级异动").toString());
            row.createCell(10).setCellValue(dbObjectList.get(i).get("异常编码").toString());
        }
        try {
            FileOutputStream fout = new FileOutputStream(path);
            wb.write(fout);
            fout.flush();
            fout.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException {
        Scanner scanner = new Scanner(System.in);
        System.out.println("此程序要完成的功能：查询并导出*对啊网*指定时期范围内的所有*订单-学员-班级*关系数据，请按以下步骤提示操作：");
        /*String start = null;
        while ((start == null) || (start.trim().isEmpty())) {
            System.out.println("(1)请输入订单支付时间-开始时间，格式yyyy-MM-dd，如：2016-01-01：");
            start = scanner.nextLine();
        }
        System.out.println(">>>>>>>>>你的输入：" + start);
        String end = null;
        while ((end == null) || (end.trim().isEmpty())) {
            System.out.println("(2)请输入订单支付时间-结束时间，格式yyyy-MM-dd，如：2016-01-31：");
            end = scanner.nextLine();
        }
        System.out.println(">>>>>>>>>你的输入：" + end);
        System.out.println("即将统计的是" + start + "~" + end + "期间的访问数量");*/
        System.out.println("(3)请输入mongo数据库IP地址，默认值：127.0.0.1，直接回车则使用默认值");
        String ip = scanner.nextLine();
        if ((ip == null) || (ip.trim().isEmpty())) {
            ip = "127.0.0.1";
        }
        System.out.println(">>>>>>>>>你的输入：" + ip);
        System.out.println("(4)请输入mongo数据库端口，默认值：27017，直接回车则使用默认值：");
        String po = scanner.nextLine();
        int port = 27017;
        if ((po != null) && (!po.trim().isEmpty())) {
            port = Integer.valueOf(po).intValue();
        }
        System.out.println(">>>>>>>>>你的输入：" + port);
        System.out.println("(5)请输入要连接的mongo数据库，默认值：local，直接回车则使用默认值：");
        String database = scanner.nextLine();
        if ((database == null) || (database.trim().isEmpty())) {
            database = "local";
        }
        System.out.println(">>>>>>>>>你的输入：" + database);
        System.out.println("(6)请输入要查询的Collection，默认值：one_one，直接回车则使用默认值：");
        String collection = scanner.nextLine();
        if ((collection == null) || (collection.trim().isEmpty())) {
            collection = "one_one";
        }
        System.out.println(">>>>>>>>>你的输入：" + collection);
        init(ip, port, database, collection);
        /*System.out.println("(7)请输入查询结果的存储路径，默认值：E:/exportDoc/，直接回车则使用默认值：");
        String idsPath = scanner.nextLine();
        if ((idsPath == null) || (idsPath.trim().isEmpty())) {
            idsPath = "E:/exportDoc/";
        }
        System.out.println(">>>>>>>>>你的输入：" + idsPath);*/
        JSONArray array = getAll();
        System.out.println("统计到的数据为：\n" + array);
        if (array == null || array.isEmpty()) {
            System.out.println("未查询到任何有效数据，表格导出失败！");
        } else {
            System.out.println("(7)请输入查询结果excel表格的保存位置，默认值根路径：E:/exportDoc/，直接回车则使用默认值：");
            String path = scanner.nextLine();
            if ((path == null) || (path.trim().isEmpty())) {
                path = "E:/exportDoc/";
            }
            System.out.println(">>>>>>>>>你的输入：" + path);
            String newFileName = null;
            while ((newFileName == null) || (newFileName.trim().isEmpty())) {
                System.out.println("(8)请输入查询结果excel表格名称（必填）：");
                newFileName = scanner.nextLine();
            }
            System.out.println(">>>>>>>>>你的输入：" + newFileName);
            exportAsExcel(array, path+newFileName);
            System.out.println("表格导出完成");
        }
        destory();
    }
}

/* Location:           D:\mongoData\jar\MongoDB-Demo.jar
 * Qualified Name:     com.xiaolong.mongo.PublicVisit
 * JD-Core Version:    0.6.0
 */