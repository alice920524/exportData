package com.myproject;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.mongodb.*;
import org.apache.poi.hssf.usermodel.*;

import java.io.*;
import java.util.*;

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

    public static List<String> readIdsFromFile(String filePath) throws IOException {
        File file = new File(filePath);
        FileInputStream fis = new FileInputStream(file);
        BufferedReader reader = new BufferedReader(new InputStreamReader(fis));
        String line = null;
        List ids = new ArrayList();
        while ((line = reader.readLine()) != null) {
            ids.add(line);
        }
        return ids;
    }

    public static JSONArray getAll(String idsPath, String startTime, String endTime) throws IOException {
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
    }

    public static void exportAsExcel(JSONArray array, String path) {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("访问记录");
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment((short)2);
        HSSFRow row0 = sheet.createRow(0);
        HSSFCell cell = row0.createCell(0);
        cell.setCellValue("订单号");
        cell.setCellStyle(style);
        for (int i = 0; i < array.size(); i++) {
            HSSFRow row = sheet.createRow(i + 1);
            VisitDetail detail = (VisitDetail)array.get(i);
            String id = detail.getId();
            List visits = detail.getVisits();
            row.createCell(0).setCellValue(id);
            for (int j = 0; j < visits.size(); j++) {
                LinkedHashMap item = (LinkedHashMap)visits.get(j);
                String date = item.get("date").toString();
                int visitNum = Integer.valueOf(item.get("visitNum").toString()).intValue();
                cell = row0.createCell(j + 1);
                cell.setCellValue(date);
                cell.setCellStyle(style);
                row.createCell(j + 1).setCellValue(visitNum);
            }
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
        System.out.println("此程序要完成的功能：统计指定时期范围内的所有公众号访问数量，公众号配置在指定路径文件中，请完成下面8个步骤");
        String start = null;
        while ((start == null) || (start.trim().isEmpty())) {
            System.out.println("(1)请输入开始日期,格式yyyyMMdd,如：20150101");
            start = scanner.nextLine();
        }
        System.out.println(">>>>>>>>>你的输入：" + start);
        String end = null;
        while ((end == null) || (end.trim().isEmpty())) {
            System.out.println("(2)请输入截止日期,格式yyyyMMdd,如：20150107");
            end = scanner.nextLine();
        }
        System.out.println(">>>>>>>>>你的输入：" + end);
        System.out.println("即将统计的是" + start + "~" + end + "期间的访问数量");
        System.out.println("(3)请输入mongo数据库IP地址,默认值：127.0.0.1,回车使用默认值");
        String ip = scanner.nextLine();
        if ((ip == null) || (ip.trim().isEmpty()))
            ip = "127.0.0.1";
        System.out.println(">>>>>>>>>你的输入：" + ip);
        System.out.println("(4)请输入mongo数据库端口,默认值：27017,回车使用默认值");
        String po = scanner.nextLine();
        int port = 27017;
        if ((po != null) && (!po.trim().isEmpty()))
            port = Integer.valueOf(po).intValue();
        System.out.println(">>>>>>>>>你的输入：" + port);
        System.out.println("(5)请输入要连接的mongo数据库,默认值：thePublic,回车使用默认值");
        String database = scanner.nextLine();
        if ((database == null) || (database.trim().isEmpty()))
            database = "thePublic";
        System.out.println(">>>>>>>>>你的输入：" + database);
        System.out.println("(6)请输入要查询的Collection,默认值：publicVisit_colletion,回车使用默认值");
        String collection = scanner.nextLine();
        if ((collection == null) || (collection.trim().isEmpty()))
            collection = "publicVisit_colletion";
        System.out.println(">>>>>>>>>你的输入：" + collection);
        init(ip, port, database, collection);
        System.out.println("(7)请输入公众号文件存储路径,默认值：E:/visit.txt,回车使用默认值");
        String idsPath = scanner.nextLine();
        if ((idsPath == null) || (idsPath.trim().isEmpty()))
            idsPath = "E:/visit.txt";
        System.out.println(">>>>>>>>>你的输入：" + idsPath);
        JSONArray array = getAll(idsPath, start, end);
        System.out.println("统计到的数据为：\n" + array);
        System.out.println("(8)请输入excel表格保存位置,默认值：E:/publicVisit.xls,回车使用默认值");
        String path = scanner.nextLine();
        if ((path == null) || (path.trim().isEmpty()))
            path = "E:/publicVisit.xls";
        System.out.println(">>>>>>>>>你的输入：" + path);
        exportAsExcel(array, path);
        System.out.println("表格导出完成");
        destory();
  }
}

/* Location:           D:\mongoData\jar\MongoDB-Demo.jar
 * Qualified Name:     com.xiaolong.mongo.PublicVisit
 * JD-Core Version:    0.6.0
 */