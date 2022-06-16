package com.chuan.chartsUtils;


import com.alibaba.fastjson.JSONArray;

import java.util.Arrays;
import java.util.List;

public class Utils {
  private static final int Row_Height_Pix = 19;   // 默认行高为19个像素
  private static final int Col_Width_Pix = 73;    // 默认列宽为73个像素

  public static List<Integer> Pix2Anchor(int width, int height, int left, int top){
    int col1 = left / Col_Width_Pix;    // left
    int row1 = top / Row_Height_Pix;   // top
    int col2 = (width + left) / Col_Width_Pix;  // width+left
    int row2 = (height + top) / Row_Height_Pix; // height+top
    return Arrays.asList(col1, row1, col2, row2);
  }

  public static String[] JsonArray2ArrayString(JSONArray jsonArray){
    String[] array = new String[jsonArray.size()];
    for(int idx = 0; idx < jsonArray.size(); idx++){
      array[idx] = jsonArray.getString(idx);
    }
    return array;
  }

  public static Double[] JsonArray2ArrayDouble(JSONArray jsonArray){
    Double[] array = new Double[jsonArray.size()];
    for(int idx = 0; idx < jsonArray.size(); idx++){
      array[idx] = jsonArray.getDouble(idx);
    }
    return array;
  }



}
