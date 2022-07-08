package com.chuan;

import com.chuan.chartsUtils.JsonParseUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;


public class Main {
  public static void main(String[] args) throws IOException {
    System.out.println("Hello world!");
    //            String jsonStr = JsonParseUtil.readJsonData("excel_sample_comment.json");
    //            String jsonStr = JsonParseUtil.readJsonData("excel_sample_pie.json");
    //        String jsonStr = JsonParseUtil.readJsonData("excel_sample_line.json");
    //        String jsonStr = JsonParseUtil.readJsonData("excel_sample.json");
    String jsonStr = JsonParseUtil.readJsonData("Excel_full.json");
    long startTime = System.currentTimeMillis();
    XSSFWorkbook excel_out = ExcelExport.exportLuckySheetByPOI(jsonStr);
    long timeProcess = System.currentTimeMillis() - startTime;
    System.out.println("处理时间：" + timeProcess);
    String date = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"));
    String fileName = "excel_out_" + date + ".xlsx";
    String Dir = "./static/excel/";
    //判断文件夹是否存在
    File parent = new File(Dir);
    if (!parent.exists()) {
      parent.mkdirs();
    }
    File file = new File(Dir, fileName);
    OutputStream os = new FileOutputStream(file);
    excel_out.write(os);
    os.flush();
    os.close();
    System.out.println("总时间：" + (System.currentTimeMillis()-startTime));
  }


}
