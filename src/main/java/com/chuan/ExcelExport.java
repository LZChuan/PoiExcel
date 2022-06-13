package com.chuan;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.awt.Color;
import java.io.*;
import java.util.*;

public class ExcelExport {
  //设置字体大小和颜色
  private static final Map<Integer, String> FontMap = new HashMap<>();

  static {
    FontMap.put(-1, "Arial");
    FontMap.put(0, "Times New Roman");
    FontMap.put(1, "Arial");
    FontMap.put(2, "Tahoma");
    FontMap.put(3, "Verdana");
    FontMap.put(4, "微软雅黑");
    FontMap.put(5, "宋体");
    FontMap.put(6, "黑体");
    FontMap.put(7, "楷体");
    FontMap.put(8, "仿宋");
    FontMap.put(9, "新宋体");
    FontMap.put(10, "华文新魏");
    FontMap.put(11, "华文行楷");
    FontMap.put(12, "华文隶书");
  }


  //基于模板导出
  public static void exportLuckySheetXlsx(String title, String newFileDir, String newFileName, String excelData) {

    JSONArray jsonArray = (JSONArray) JSONObject.parse(excelData);
    JSONObject jsonObject = (JSONObject) jsonArray.get(0);
    JSONArray jsonObjectList = jsonObject.getJSONArray("celldata");

    //excel模板路径
    String filePath = "file/" + "模板.xlsx";
    //        String filePath = "/Users/ouyang/Downloads/uploadTestProductFile/生产日报表.xlsx";
    File file = new File(filePath);
    FileInputStream in = null;
    try {
      in = new FileInputStream(file);
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    }
    //读取excel模板
    XSSFWorkbook wb = null;
    try {
      wb = new XSSFWorkbook(in);
    } catch (IOException e) {
      e.printStackTrace();
    }
    //读取了模板内所有sheet内容
    XSSFSheet sheet = wb.getSheetAt(0);
    //如果这行没有了，整个公式都不会有自动计算的效果的
    sheet.setForceFormulaRecalculation(true);

    for (int index = 0; index < jsonObjectList.size(); index++) {
      JSONObject object = jsonObjectList.getJSONObject(index);
      String str_ = object.get("r") + "_" + object.get("c") + "=" + ((JSONObject) object.get("v")).get("v") + "\n";
      JSONObject jsonObjectValue = ((JSONObject) object.get("v"));

      String value = "";
      if (jsonObjectValue != null && jsonObjectValue.get("v") != null) {
        value = jsonObjectValue.get("v") + "";
      }
      if (sheet.getRow((int) object.get("r")) != null && sheet.getRow((int) object.get("r")).getCell((int) object.get("c")) != null) {
        sheet.getRow((int) object.get("r")).getCell((int) object.get("c")).setCellValue(value);
      } else {
        System.out.println("错误的=" + index + ">>>" + str_);
      }
    }

    // 保存文件的路径
    //        String realPath = "/Users/ouyang/Downloads/uploadTestProductFile/其他文件列表/生产日报表/";
    // 判断路径是否存在
    File dir = new File(newFileDir);
    if (!dir.exists()) {
      dir.mkdirs();
    }
    //修改模板内容导出新模板
    FileOutputStream out = null;
    try {
      out = new FileOutputStream(newFileDir + newFileName);
      wb.write(out);
      out.close();
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } catch (IOException e) {
      e.printStackTrace();
    }
    System.out.println("生成文件成功：" + newFileDir + newFileName);
  }

  /***
   * 基于POI解析 从0开始导出xlsx文件，不是基于模板
   * @param title 表格名
   * @param newFileDir 保存的文件夹名
   * @param newFileName 保存的文件名
   * @param excelData luckysheet 表格数据
   */
  public static void exportLuckySheetXlsxByPOI(String title, String newFileDir, String newFileName, String excelData) {
    excelData = excelData.replace("&#xA;", "\\r\\n");//去除luckysheet中 &#xA 的换行

    JSONArray jsonArray = (JSONArray) JSONObject.parse(excelData);


    for (int sheetIndex = 0; sheetIndex < jsonArray.size(); sheetIndex++) {

      JSONObject jsonObject = (JSONObject) jsonArray.get(sheetIndex);

      JSONArray celldataObjectList = jsonObject.getJSONArray("celldata");

      JSONArray rowObjectList = jsonObject.getJSONArray("visibledatarow");

      JSONArray colObjectList = jsonObject.getJSONArray("visibledatacolumn");

      JSONArray dataObjectList = jsonObject.getJSONArray("data");
      JSONObject mergeObject = jsonObject.getJSONObject("config").getJSONObject("merge");//合并单元格
      JSONObject columnlenObject = jsonObject.getJSONObject("config").getJSONObject("columnlen");//表格列宽
      JSONObject rowlenObject = jsonObject.getJSONObject("config").getJSONObject("rowlen");//表格行高
      JSONArray borderInfoObjectList = jsonObject.getJSONObject("config").getJSONArray("borderInfo");//边框样式
      //参考：https://blog.csdn.net/jdtugfcg/article/details/84100315
      //创建操作Excel的XSSFWorkbook对象
      XSSFWorkbook excel = new XSSFWorkbook();
      XSSFCellStyle cellStyle = excel.createCellStyle();
      //创建XSSFSheet对象
      XSSFSheet sheet = excel.createSheet(jsonObject.getString("name"));

      //我们都知道excel是表格，即由一行一行组成的，那么这一行在java类中就是一个XSSFRow对象，我们通过XSSFSheet对象就可以创建XSSFRow对象
      //如：创建表格中的第一行（我们常用来做标题的行)  XSSFRow firstRow = sheet.createRow(0); 注意下标从0开始
      //根据luckysheet创建行列
      //创建行和列
      for (int i = 0; i < rowObjectList.size(); i++) {
        XSSFRow row = sheet.createRow(i);//创建行
        try {
          row.setHeightInPoints(Float.parseFloat(rowlenObject.get(i) + ""));//行高px值
        } catch (Exception e) {
          row.setHeightInPoints(20f);//默认行高
        }

        for (int j = 0; j < colObjectList.size(); j++) {
          if (columnlenObject.getInteger(j + "") != null) {
            sheet.setColumnWidth(j, columnlenObject.getInteger(j + "") * 42);//列宽px值
          }
          row.createCell(j);//创建列
        }
      }

      //设置值,样式
      setCellValue(celldataObjectList, borderInfoObjectList, sheet, excel);

      // 判断路径是否存在
      File dir = new File(newFileDir);
      if (!dir.exists()) {
        dir.mkdirs();
      }
      OutputStream out = null;
      try {
        out = new FileOutputStream(newFileDir + newFileName);

        excel.write(out);

        out.close();

      } catch (FileNotFoundException e) {
        e.printStackTrace();

      } catch (IOException e) {
        e.printStackTrace();

      }
    }


  }


  /***
   * 导出excel通过POI实现
   * @param excelData 前端数据表格的json
   */
  public static XSSFWorkbook exportLuckySheetByPOI(String excelData) {
    //去除luckysheet中 &#xA 的换行
    excelData = excelData.replace("&#xA;", "\\r\\n");
    JSONArray jsonArray = JsonParseUtil.parseStrToJson(excelData);
    //创建操作Excel的XSSFWorkbook对象
    XSSFWorkbook excel = new XSSFWorkbook();

    //单元格样式，一个对象重复使用
    //    CellStyle cellStyle = excel.createCellStyle();

    //有多少个表就循环多少次
    for (int sheetIndex = 0; sheetIndex < jsonArray.size(); sheetIndex++) {
      //获取配置
      JSONObject jsonObject = jsonArray.getJSONObject(sheetIndex);
      //初始化的数据
      JSONArray celldata = jsonObject.getJSONArray("celldata");
      //所有行的位置
      JSONArray visibledatarow = jsonObject.getJSONArray("visibledatarow");
      //所有类的位置
      JSONArray visibledatacolumn = jsonObject.getJSONArray("visibledatacolumn");
      //更新和存储使用的单元格数据
      JSONArray data = jsonObject.getJSONArray("data");
      // 图表数据
      JSONArray chart = jsonObject.getJSONArray("chart");
      // 表的整体配置
      JSONObject config = jsonObject.getJSONObject("config");

      //单元格的样式
      XSSFCellStyle cellStyle = excel.createCellStyle();
      //创建XSSFSheet对象并命名
      XSSFSheet sheet = excel.createSheet(jsonObject.getString("name"));
      //创建表格
      createRowsAndColumns(excel, sheet, data, config, celldata);
    }

    return excel;

  }

  /**
   * 构造表格结构，先把每个表的行和列创建出来
   *
   * @param sheet  工作表
   * @param data   工作表数据
   * @param config 工作表配置
   */
  private static void createRowsAndColumns(XSSFWorkbook excel, XSSFSheet sheet, JSONArray data, JSONObject config, JSONArray cellData) {
    /***
     * luckysheet的配置，config下的配置
     */
    /************************************这里准备几个对象用来存放luckysheet的config中的配置*************************************/
    //行高
    JSONObject rowlen = null;
    //列宽
    JSONObject columnlen = null;
    //边框信息
    JSONArray borderInfo = null;
    //合并单元格信息
    JSONObject merge = null;
    /************************************这里准备几个对象用来存放luckysheet的config中的配置*************************************/
    rowlen = config.getJSONObject("rowlen");
    columnlen = config.getJSONObject("columnlen");
    borderInfo = config.getJSONArray("borderInfo");

    //行
    XSSFRow row = null;
    for (int i = 0; i < data.size(); i++) {
      row = sheet.createRow(i);
      //设置行高,默认为20
      if (rowlen.size() > 0) {
        row.setHeightInPoints(rowlen.getInteger(String.valueOf(i)) != null ? rowlen.getInteger(String.valueOf(i)) : 20);
      }

      //创建列
      for (int j = 0; j < data.getJSONArray(i).size(); j++) {
        // 设置列宽，无默认值  TODO 能否在外层设置？
        if (columnlen != null && columnlen.getInteger(String.valueOf(j)) != null) {
          sheet.setColumnWidth(j, columnlen.getInteger(j + "") * 42);//列宽px值
        }
        //这里可以设置celltype，在构造方法里

        // 创建单元格
        row.createCell(j);
      }
    }

    //设置值单元格值
    setCellValue(cellData, borderInfo, sheet, excel);
  }

  private static void setCellValue(JSONArray cellData, JSONArray borderInfoObjectList, XSSFSheet
      sheet, XSSFWorkbook excel) {
    String cellType = "";
    String cellFormat = "";

    // 设置所有单元格信息
    for (int index = 0; index < cellData.size(); index++) {

      XSSFCellStyle style = excel.createCellStyle();//样式
      XSSFFont font = excel.createFont();//字体样式
      //数字格式
      XSSFDataFormat dataFormat = excel.createDataFormat();

      JSONObject object = cellData.getJSONObject(index);

      //单元格类型
      if (object.getJSONObject("v").containsKey("ct")) {
        // Type类型
        cellType = object.getJSONObject("v").getJSONObject("ct").getString("t");
        // Format格式的定义串
        cellFormat = object.getJSONObject("v").getJSONObject("ct").getString("fa");
      }

      String str_ = object.get("r") + "_" + object.get("c") + "=" + ((JSONObject) object.get("v")).get("v") + "\n";
      JSONObject jsonObjectValue = ((JSONObject) object.get("v"));

      String value = "";
      if (jsonObjectValue != null && jsonObjectValue.get("v") != null) {
        value = jsonObjectValue.getString("v");
      }

      if (sheet.getRow((int) object.get("r")) != null && sheet.getRow((int) object.get("r")).getCell((int) object.get("c")) != null) {
        XSSFCell cell = sheet.getRow((int) object.get("r")).getCell((int) object.get("c"));

        if (jsonObjectValue != null && jsonObjectValue.get("f") != null) {//如果有公式，设置公式
          value = jsonObjectValue.getString("f");
          cell.setCellFormula(value.substring(1, value.length()));//不需要=符号
        }

        if (cellFormat != null || !cellFormat.equals("")) {
          style.setDataFormat(dataFormat.getFormat(cellFormat));
        }

        //合并单元格与填充单元格颜色
        setMergeAndColorByObject(jsonObjectValue, sheet, style);
        JSONObject mergeObject = (JSONObject) jsonObjectValue.get("mc");
        if (mergeObject != null) {
          int r = (int) (mergeObject.get("r"));
          int c = (int) (mergeObject.get("c"));
          if ((mergeObject.get("rs") != null && (mergeObject.get("cs") != null))) {
            int rs = (int) (mergeObject.get("rs"));
            int cs = (int) (mergeObject.get("cs"));
            CellRangeAddress region = new CellRangeAddress(r, r + rs - 1, (short) (c), (short) (c + cs - 1));
            sheet.addMergedRegion(region);
          }
        }

        if (jsonObjectValue.getString("bg") != null) {
          int bg = Integer.parseInt(jsonObjectValue.getString("bg").replace("#", ""), 16);
          style.setFillPattern(FillPatternType.SOLID_FOREGROUND);    //设置填充方案
          style.setFillForegroundColor(new XSSFColor((IndexedColorMap) new Color(bg)));  //设置填充颜色
        }

        if (!value.equals("")) {
          style.setDataFormat(dataFormat.getFormat(cellFormat));
          if (cellType == "n") {
            cell.setCellValue(Double.parseDouble(value));
          } else {
            cell.setCellValue(value);
          }

          //          //处理单元格格式，主要处理数字，double类型和百分比,如果是数字
          //          if(NUMBER_PATTERN.matcher(value).matches()){ //如果是数字
          ////          if(true){ //如果是数字
          //            style.setDataFormat(dataFormat.getFormat(cellFormat));
          //            cell.setCellValue(Double.parseDouble(value));
          //          }else{   //其他格式
          //            style.setDataFormat(dataFormat.getFormat(cellFormat));
          //            cell.setCellValue(value);
          //          }
        }

        //设置垂直水平对齐方式
        int vt = jsonObjectValue.getInteger("vt") == null ? 1 : jsonObjectValue.getInteger("vt");//垂直对齐	 0 中间、1 上、2下
        int ht = jsonObjectValue.getInteger("ht") == null ? 1 : jsonObjectValue.getInteger("ht");//0 居中、1 左、2右
        switch (vt) {
          case 0:
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            break;
          case 1:
            style.setVerticalAlignment(VerticalAlignment.TOP);
            break;
          case 2:
            style.setVerticalAlignment(VerticalAlignment.BOTTOM);
            break;
        }
        switch (ht) {
          case 0:
            style.setAlignment(HorizontalAlignment.CENTER);
            break;
          case 1:
            style.setAlignment(HorizontalAlignment.LEFT);
            break;
          case 2:
            style.setAlignment(HorizontalAlignment.RIGHT);
            break;
        }

        //设置合并单元格的样式有问题
        String ff = jsonObjectValue.getString("ff");//0 Times New Roman、 1 Arial、2 Tahoma 、3 Verdana、4 微软雅黑、5 宋体（Song）、6 黑体（ST Heiti）、7 楷体（ST Kaiti）、 8 仿宋（ST FangSong）、9 新宋体（ST Song）、10 华文新魏、11 华文行楷、12 华文隶书
        int fs = jsonObjectValue.getInteger("fs") == null ? 14 : jsonObjectValue.getInteger("fs");//字体大小
        int bl = jsonObjectValue.getInteger("bl") == null ? 0 : jsonObjectValue.getInteger("bl");//粗体	0 常规 、 1加粗
        int it = jsonObjectValue.getInteger("it") == null ? 0 : jsonObjectValue.getInteger("it");//斜体	0 常规 、 1 斜体
        String fc = jsonObjectValue.getString("fc") == null ? "" : jsonObjectValue.getString("fc");//字体颜色
        font.setFontName(FontMap.get(ff));//字体名字
        // rgb颜色转换
        String[] rgb_str = fc.replace("rgb(", "").replace(")", "").replace(" ", "").split(",");
        if (fc.length() > 0) {
          font.setColor(new XSSFColor(
              new Color(Integer.parseInt(rgb_str[0]), Integer.parseInt(rgb_str[1]), Integer.parseInt(rgb_str[2])),
              new DefaultIndexedColorMap()));
        }
        font.setFontName(ff);//字体名字
        font.setFontHeightInPoints((short) fs);//字体大小
        if (bl == 1) {
          font.setBold(true);//粗体显示
        }
        font.setItalic(it == 1 ? true : false);//斜体
        style.setFont(font);
        style.setWrapText(true);//设置自动换行
        cell.setCellStyle(style);

      } else {
        System.out.println("错误的=" + index + ">>>" + str_);
      }
    }
    //设置边框
    setBorder(borderInfoObjectList, excel, sheet);
  }

  //设置边框
  private static void setBorder(JSONArray borderInfoObjectList, XSSFWorkbook workbook, XSSFSheet sheet) {
    //设置边框样式map
    Map<Integer, BorderStyle> bordMap = new HashMap<>();
    bordMap.put(1, BorderStyle.THIN);
    bordMap.put(2, BorderStyle.HAIR);
    bordMap.put(3, BorderStyle.DOTTED);
    bordMap.put(4, BorderStyle.DASHED);
    bordMap.put(5, BorderStyle.DASH_DOT);
    bordMap.put(6, BorderStyle.DASH_DOT_DOT);
    bordMap.put(7, BorderStyle.DOUBLE);
    bordMap.put(8, BorderStyle.MEDIUM);
    bordMap.put(9, BorderStyle.MEDIUM_DASHED);
    bordMap.put(10, BorderStyle.MEDIUM_DASH_DOT);
    bordMap.put(11, BorderStyle.MEDIUM_DASH_DOT_DOT);
    bordMap.put(12, BorderStyle.SLANTED_DASH_DOT);
    bordMap.put(13, BorderStyle.THICK);

    //一定要通过 cell.getCellStyle()  不然的话之前设置的样式会丢失
    //设置边框
    if (null != borderInfoObjectList) {
      for (int i = 0; i < borderInfoObjectList.size(); i++) {
        JSONObject borderInfoObject = (JSONObject) borderInfoObjectList.get(i);
        if (borderInfoObject.get("rangeType").equals("cell")) {//单个单元格
          JSONObject borderValueObject = borderInfoObject.getJSONObject("value");

          JSONObject l = borderValueObject.getJSONObject("l");
          JSONObject r = borderValueObject.getJSONObject("r");
          JSONObject t = borderValueObject.getJSONObject("t");
          JSONObject b = borderValueObject.getJSONObject("b");


          int row = borderValueObject.getInteger("row_index");
          int col = borderValueObject.getInteger("col_index");

          XSSFCell cell = sheet.getRow(row).getCell(col);


          if (l != null) {
            cell.getCellStyle().setBorderLeft(bordMap.get((int) l.get("style"))); //左边框
            int bg = Integer.parseInt(l.getString("color").replace("#", ""), 16);
            cell.getCellStyle().setLeftBorderColor(new XSSFColor(new Color(bg)));//左边框颜色
          }
          if (r != null) {
            cell.getCellStyle().setBorderRight(bordMap.get((int) r.get("style"))); //右边框
            int bg = Integer.parseInt(r.getString("color").replace("#", ""), 16);
            cell.getCellStyle().setRightBorderColor(new XSSFColor(new Color(bg)));//右边框颜色
          }
          if (t != null) {
            cell.getCellStyle().setBorderTop(bordMap.get((int) t.get("style"))); //顶部边框
            int bg = Integer.parseInt(t.getString("color").replace("#", ""), 16);
            cell.getCellStyle().setTopBorderColor(new XSSFColor(new Color(bg)));//顶部边框颜色
          }
          if (b != null) {
            cell.getCellStyle().setBorderBottom(bordMap.get((int) b.get("style"))); //底部边框
            int bg = Integer.parseInt(b.getString("color").replace("#", ""), 16);
            cell.getCellStyle().setBottomBorderColor(new XSSFColor(new Color(bg)));//底部边框颜色
          }
        } else if (borderInfoObject.get("rangeType").equals("range")) {//选区
          int bg_ = Integer.parseInt(borderInfoObject.getString("color").replace("#", ""), 16);
          int style_ = borderInfoObject.getInteger("style");

          JSONObject rangObject = (JSONObject) ((JSONArray) (borderInfoObject.get("range"))).get(0);

          JSONArray rowList = rangObject.getJSONArray("row");
          JSONArray columnList = rangObject.getJSONArray("column");


          for (int row_ = rowList.getInteger(0); row_ < rowList.getInteger(rowList.size() - 1) + 1; row_++) {
            for (int col_ = columnList.getInteger(0); col_ < columnList.getInteger(columnList.size() - 1) + 1; col_++) {
              XSSFCell cell = sheet.getRow(row_).getCell(col_);

              cell.getCellStyle().setBorderLeft(bordMap.get(style_)); //左边框
              cell.getCellStyle().setLeftBorderColor(new XSSFColor(new Color(bg_)));//左边框颜色
              cell.getCellStyle().setBorderRight(bordMap.get(style_)); //右边框
              cell.getCellStyle().setRightBorderColor(new XSSFColor(new Color(bg_)));//右边框颜色
              cell.getCellStyle().setBorderTop(bordMap.get(style_)); //顶部边框
              cell.getCellStyle().setTopBorderColor(new XSSFColor(new Color(bg_)));//顶部边框颜色
              cell.getCellStyle().setBorderBottom(bordMap.get(style_)); //底部边框
              cell.getCellStyle().setBottomBorderColor(new XSSFColor(new Color(bg_)));//底部边框颜色 }
            }
          }
        }
      }
    }

  }

  //设置合并单元格
  private static void setMergeAndColorByObject(JSONObject jsonObjectValue, XSSFSheet sheet, XSSFCellStyle style) {
    JSONObject mergeObject = (JSONObject) jsonObjectValue.get("mc");
    if (mergeObject != null) {
      int r = (int) (mergeObject.get("r"));
      int c = (int) (mergeObject.get("c"));
      if ((mergeObject.get("rs") != null && (mergeObject.get("cs") != null))) {
        int rs = (int) (mergeObject.get("rs"));
        int cs = (int) (mergeObject.get("cs"));
        CellRangeAddress region = new CellRangeAddress(r, r + rs - 1, (short) (c), (short) (c + cs - 1));
        sheet.addMergedRegion(region);
      }
    }

    if (jsonObjectValue.getString("bg") != null) {
      int bg = Integer.parseInt(jsonObjectValue.getString("bg").replace("#", ""), 16);
      style.setFillPattern(FillPatternType.SOLID_FOREGROUND);    //设置填充方案
      style.setFillForegroundColor(new XSSFColor((IndexedColorMap) new Color(bg)));  //设置填充颜色
    }
  }


}
