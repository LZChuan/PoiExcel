package com.chuan;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;


public class Main {
    public static <PortableConnectionImpl> void main(String[] args) throws IOException {
        System.out.println("Hello world!");
        String jsonStr = JsonParseUtil.readJsonData("excel_sample.json");
//        String jsonStr = JsonParseUtil.readJsonData("Excel_full.json");
        XSSFWorkbook excel_out = ExcelExport.exportLuckySheetByPOI(jsonStr);
//        System.out.println(jsonStr);

        String date = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"));
        String fileName="excel_out_"+date+".xlsx";

        String Dir = "./static/excel/";
        //判断文件夹是否存在
        File parent = new File(Dir);
        if (!parent.exists()) {
            parent.mkdirs();
        }
        File file=new File(Dir,fileName);
        OutputStream os=new FileOutputStream(file);
        excel_out.write(os);
        os.flush();
        os.close();
    }


}