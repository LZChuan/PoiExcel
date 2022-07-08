package com.chuan.chartsUtils;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;

import java.io.*;
import java.nio.charset.StandardCharsets;

public class JsonParseUtil {

  public static String readJsonData(String pactFile) throws IOException {
    // 读取文件数据
    //System.out.println("读取文件数据util");

    StringBuilder strbuffer = new StringBuilder();
    File myFile = new File(pactFile);//"D:"+File.separatorChar+"DStores.json"
    if (!myFile.exists()) {
      System.err.println("Can't Find " + pactFile);
    }
    try {
      FileInputStream fis = new FileInputStream(pactFile);
      InputStreamReader inputStreamReader = new InputStreamReader(fis, StandardCharsets.UTF_8);
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
