package com.myproject.common;

/**
 * Created by admin on 2017/11/24.
 */

import com.mongodb.DBCollection;
import com.mongodb.DBObject;
import com.mongodb.client.MongoCollection;
import com.mongodb.util.JSON;
import com.myproject.utils.MongoDbUtils;
import org.bson.Document;

/**
 * 对于MongoDB 数据库操作的方法
 */
public class C_MONGODB {

    /**
     * 将数据存储到MongoDB中。
     *
     * @param json 要存储到MongoDB数据库的字符串
     * @param collection MongoDB中集合的名称（表名）
     * @return 插入数据的主键
     * @throws Exception
     */
    public static String saveObjectByJson(String dbName, String json, String collection) throws Exception {
        MongoCollection<Document> coll = MongoDbUtils.getCollection(dbName, collection);
        DBObject dbobject = (DBObject) JSON.parse(json);
//        coll.save(dbobject);
        String oid = dbobject.get("_id").toString();

        return oid;
    }
}
