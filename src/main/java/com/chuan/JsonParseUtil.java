package com.chuan;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

public class JsonParseUtil {

  //    将存入的配置转为jsonArray格式
  public static JSONArray parseStrToJson(String option){
    StringBuffer sb = new StringBuffer("{'':");
    sb.append(option);
    sb.append("}");

    //        转换为json格式
    JSONObject jsonObject = JSONObject.parseObject(sb.toString());

    //        转换为array
    JSONArray jsonArray = jsonObject.getJSONArray("");

    return jsonArray;
  }

}
