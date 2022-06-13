package com.chuan;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

import java.io.*;

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
  public static String readJsonData(String pactFile) throws IOException {
    // 读取文件数据
    //System.out.println("读取文件数据util");

    StringBuffer strbuffer = new StringBuffer();
    File myFile = new File(pactFile);//"D:"+File.separatorChar+"DStores.json"
    if (!myFile.exists()) {
      System.err.println("Can't Find " + pactFile);
    }
    try {
      FileInputStream fis = new FileInputStream(pactFile);
      InputStreamReader inputStreamReader = new InputStreamReader(fis, "UTF-8");
      BufferedReader in  = new BufferedReader(inputStreamReader);

      String str;
      while ((str = in.readLine()) != null) {
        strbuffer.append(str);  //new String(str,"UTF-8")
      }
      in.close();
    } catch (IOException e) {
      e.getStackTrace();
    }
    //System.out.println("读取文件结束util");
    return strbuffer.toString();
  }

}
