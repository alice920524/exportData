package com.myproject.web;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.myproject.entity.RequestData;
import com.myproject.utils.MongoDbUtils;
import org.bson.Document;
import org.bson.conversions.Bson;
import org.springframework.stereotype.Controller;
import org.springframework.ui.ModelMap;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.RequestMapping;

import com.myproject.entity.RequestData;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.model.Filters;
import com.myproject.utils.MongoDbUtils;
import com.myproject.dao.MongoDBDao;

/**
 * @Title MongdbController.java
 * @Description TODO
 * @author night
 * @date 2015年12月29日 下午3:39:37
 * @version V1.0
 */
@Controller
public class MongoDBController {
    private int pageSize = 2;

    @RequestMapping(value = "/add")
    public String add(@ModelAttribute RequestData data){
//       Map<String, String> map = System.getenv();
//          for(Iterator<String> itr = map.keySet().iterator();itr.hasNext();){
//              String key = itr.next();
//              System.out.println(key + "=" + map.get(key));
//          }
//
//      System.out.println(">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
//      Properties p =  System.getProperties();
//      for(Iterator<Object> itr = p.keySet().iterator();itr.hasNext();){
//            String key = (String) itr.next();
//            System.out.println(key + "=" + p.get(key));
//        }
        data.setFirstDate(new Date());
        Document document = data.toDocument(data);
        MongoDbUtils.insertOne("student", document);
        return "redirect:/queryAll";
    }
    @RequestMapping(value = "/queryAll")
    public String queryAll(ModelMap modelMap){
//一种查询方式
//      Document filter =new Document();
//      filter.append(Seach.NE.getOperStr(), "-12");
//      Document seachDocument=new Document();
//      seachDocument.put("_id", filter);

        //另一种查询方式
        //Bson filter = Filters.ne("_id", "-12");不等于
        //Bson filter = Filters.eq("_id", "-12");等于
        //查询卡号分别为 1，3 ，4的数据
        //String[] items = {"1","3","4"};
        //Bson filter = Filters.in("cardNo", items);
        //查询当前五分钟之内的数据   firstDate在mongodb中是直接以Date类型存储的
        Date date = new Date();
        date.setMinutes(-5);
        System.out.println(date.toLocaleString());
        Bson filter = Filters.and(Filters.lte("firstDate", new Date()),Filters.gte("firstDate", date));
        MongoCursor<Document> result = MongoDbUtils.findByFilter("student", filter);
        List<RequestData> listData = new ArrayList<RequestData>();
        int count = 0;
        System.out.println("--》第" + count +"条");
        while (result.hasNext()) {
            count++;
            System.out.println("第" + count +"条");
            Document document=result.next();
            RequestData requestData=new RequestData();
            requestData.toRequestData(document);
            listData.add(requestData);
        }
        modelMap.put("listData", listData);
        return "result";
    }

    @RequestMapping(value = "/del")
    public String del(String id){
        System.out.println("需要删除数据id:" + id);


        int delCount  = MongoDbUtils.deleteById("student", id);
        System.out.println("删除条数：" + delCount);
        return "redirect:/queryAll";
    }

    @RequestMapping(value="/update")
    public String update(String id,ModelMap model){
        System.out.println("要修改的数据id：" + id);
        Document document = MongoDbUtils.findById("student", id);
        RequestData data = new RequestData();
        data.toRequestData(document);
        model.put("data", data);
        return "update";
    }

    @RequestMapping(value="/modify")
    public String modify(@ModelAttribute RequestData data,ModelMap model){
        System.out.println("要修改的数据id：" + data.getId());
        data.setFirstDate(new Date());
        Document newDocument = data.toDocument(data);
        int updateCount = MongoDbUtils.updateById("student", data.getId().toString(), newDocument);
        System.out.println("修改条数：" + updateCount);
        return "redirect:/queryAll";
    }

    @RequestMapping(value = "/query")
    public String query(int pageNo,ModelMap modelMap){
        System.out.println("去的页数：" + pageNo);
        Bson filter = Filters.ne("_id", "-12");
        MongoCursor<Document> result = MongoDbUtils.findByPage("student", filter, null, pageNo, pageSize);
        List<RequestData> listData = new ArrayList<RequestData>();
        int count = 0;
        System.out.println("--》第" + count +"条");
        while (result.hasNext()) {
            count++;
            System.out.println("第" + count +"条");
            Document document=result.next();
            RequestData requestData=new RequestData();
            requestData.toRequestData(document);
            listData.add(requestData);
        }
        modelMap.put("listData", listData);
        return "result";
    }

}
