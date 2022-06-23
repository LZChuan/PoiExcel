package com.chuan;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import jdk.nashorn.internal.runtime.regexp.joni.exception.ValueException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.*;
import com.chuan.chartsUtils.*;

import java.awt.Color;
import java.io.*;
import java.util.*;

public class ExcelExport {
  //设置字体大小和颜色
  private static final Map<Integer, String> FontMap = new HashMap<>();
  //设置边框样式map
  private static final Map<Integer, BorderStyle> BordMap = new HashMap<>();
  //设置默认颜色库
  private static final DefaultIndexedColorMap DEFAULT_INDEXED_COLOR_MAP = new DefaultIndexedColorMap();

  static {
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

    BordMap.put(1, BorderStyle.THIN);
    BordMap.put(2, BorderStyle.HAIR);
    BordMap.put(3, BorderStyle.DOTTED);
    BordMap.put(4, BorderStyle.DASHED);
    BordMap.put(5, BorderStyle.DASH_DOT);
    BordMap.put(6, BorderStyle.DASH_DOT_DOT);
    BordMap.put(7, BorderStyle.DOUBLE);
    BordMap.put(8, BorderStyle.MEDIUM);
    BordMap.put(9, BorderStyle.MEDIUM_DASHED);
    BordMap.put(10, BorderStyle.MEDIUM_DASH_DOT);
    BordMap.put(11, BorderStyle.MEDIUM_DASH_DOT_DOT);
    BordMap.put(12, BorderStyle.SLANTED_DASH_DOT);
    BordMap.put(13, BorderStyle.THICK);
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
      JSONArray chartData = jsonObject.getJSONArray("chart");
      // 表的整体配置
      JSONObject config = jsonObject.getJSONObject("config");

      boolean isPivotTable = jsonObject.getBoolean("isPivotTable") != null && jsonObject.getBoolean("isPivotTable");
      if (isPivotTable) {   // TODO 透视表暂时不支持
        continue;
      }
      //单元格的样式
      //创建XSSFSheet对象并命名
      XSSFSheet sheet = excel.createSheet(jsonObject.getString("name"));
      // 给sheet填充数据
      sheet = createRowsAndColumns(excel, sheet, data, config, celldata);
      // 冻结
      if (jsonObject.containsKey("frozen")) {
        String frozenType = jsonObject.getJSONObject("frozen").getString("type");
        int row_focus, column_focus;
        JSONObject frozenRange;
        switch (frozenType) {
          case "row": // 首行冻结
            sheet.createFreezePane(0, 1);
            break;
          case "column": // 冻结首列
            sheet.createFreezePane(1, 0);
            break;
          case "both": // 冻结行列
            sheet.createFreezePane(1, 1);
            break;
          case "rangeRow": // 冻结行到选区
          case "rangeColumn": // 冻结列到选区
          case "rangeBoth": // 冻结行列到选区
            frozenRange = jsonObject.getJSONObject("frozen").getJSONObject("range");
            row_focus = frozenRange.getInteger("row_focus");
            column_focus = frozenRange.getInteger("column_focus");
            sheet.createFreezePane(column_focus, row_focus);
            break;
          case "cancel": // 取消冻结
            break;
        }
      }

      // 给 sheet绘图
      if (chartData != null && chartData.size() > 0) {
        DrawSheetCharts.SheetCharts(sheet, chartData);
      }
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
  private static XSSFSheet createRowsAndColumns(XSSFWorkbook excel, XSSFSheet sheet, JSONArray data, JSONObject config, JSONArray cellData) {
    /***
     * luckysheet的配置，config下的配置
     */
    /************************************这里准备几个对象用来存放luckysheet的config中的配置*************************************/
    //行高
    JSONObject rowlen = config.getJSONObject("rowlen");
    //列宽
    JSONObject columnlen = config.getJSONObject("columnlen");
    //边框信息
    JSONArray borderInfo = config.getJSONArray("borderInfo");
    //隐藏行
    JSONObject rowHidden = config.getJSONObject("rowhidden");

    /************************************这里准备几个对象用来存放luckysheet的config中的配置*************************************/
    // 创建 行
    XSSFRow row = null;
    for (int i = 0; i < data.size(); i++) {
      // 创建行并设置行高
      row = sheet.createRow(i);
      // ######################### 设置行高,默认为20  ################
      float rowhigh = 20f;
      if (rowlen != null && rowlen.getInteger(String.valueOf(i)) != null) {
        rowhigh = rowlen.getFloat(String.valueOf(i));
      }
      if (rowHidden != null && rowHidden.getInteger(String.valueOf(i)) != null) {
        rowhigh = 0;
      }
      row.setHeightInPoints(rowhigh);

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
    //设置所有单元格值
    XSSFSheet this_sheet = setCellValue(cellData, borderInfo, sheet, excel);

    return this_sheet;
  }

  private static XSSFSheet setCellValue(JSONArray cellData, JSONArray borderInfoObjectList, XSSFSheet sheet, XSSFWorkbook excel) {
    String cellType = "";
    String cellFormat = "";

    // 设置所有单元格信息
    for (int index = 0; index < cellData.size(); index++) {

      XSSFCellStyle style = excel.createCellStyle();   // TODO 反复创建，拖慢速度，需要优化
      XSSFFont font = excel.createFont();//字体样式
      //数字格式
      XSSFDataFormat dataFormat = excel.createDataFormat();
      JSONObject cellObject = cellData.getJSONObject(index);

      if (sheet.getRow((int) cellObject.get("r")) != null
          && sheet.getRow((int) cellObject.get("r")).getCell((int) cellObject.get("c")) != null) {
        XSSFCell cell = sheet.getRow((int) cellObject.get("r")).getCell((int) cellObject.get("c"));

        if (cellObject.getJSONObject("v") != null) {
          // 取单元格内容数据 Object
          JSONObject cellObject_v = cellObject.getJSONObject("v");
          //单元格内容 类型
          if (cellObject_v.containsKey("ct")) {
            // Type类型
            cellType = cellObject_v.getJSONObject("ct").getString("t");
            // Format格式的定义串
            cellFormat = cellObject_v.getJSONObject("ct").getString("fa");
            if (cellFormat != null && !Objects.equals(cellFormat, "")) {  // TODO 需要验证条件
              style.setDataFormat(dataFormat.getFormat(cellFormat));
            }
          }

          // 单元格内容填充
          String cellValue;
          //如果有公式，则使用公式
          if (cellObject_v.get("f") != null) {
            // TODO 有很多公式无法支持，使用try进行设置
            try {
              cellValue = cellObject_v.getString("f");
              cell.setCellFormula(cellValue.substring(1));
            } catch (Exception e) {
              e.printStackTrace();
            }
          } else if (cellObject_v.get("v") != null) {
            // 没有公式则用“v”中的值
            cellValue = cellObject_v.getString("v");
            // 根据单元格 类型 修改单元格内容的值  TODO 需要完善，识别 各种类型；
            if (!cellValue.equals("")) {
              switch (cellType) {
                case "s": // 纯文本
                case "g": // 默认格式
                  cell.setCellValue(cellValue);
                  break;
                case "d": // 时间
                  cell.setCellValue(Double.parseDouble(cellValue));
                  break;
                case "n":
                  try {
                    cell.setCellValue(Double.parseDouble(cellValue));
                  } catch (Exception e) {
                    e.printStackTrace();
                    // 会出现 公式的情况
                    cell.setCellValue(cellValue);
                  }
                  break;
              }
            }
          }

          //合并单元格
          if (cellObject_v.containsKey("mc")) {
            JSONObject mergeObject = (JSONObject) cellObject_v.get("mc");
            if (mergeObject != null) {
              int r = mergeObject.getInteger("r");
              int c = mergeObject.getInteger("c");
              if (mergeObject.get("rs") != null && (mergeObject.get("cs") != null)) {
                int rs = mergeObject.getInteger("rs");
                int cs = mergeObject.getInteger("cs");
                CellRangeAddress region = new CellRangeAddress(r, r + rs - 1, (short) (c), (short) (c + cs - 1));
                sheet.addMergedRegion(region);
              }
            }
          }

          // 单元格背景颜色
          if (cellObject_v.getString("bg") != null) {
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);    //设置填充方案
            style.setFillForegroundColor(new XSSFColor(setColor(cellObject_v.getString("bg")), new DefaultIndexedColorMap()));
          }

          //************************* 字体设置 ************************
          // 设置字体是否加粗 	0 常规, 1加粗
          font.setBold(cellObject_v.getBoolean("bl") != null && cellObject_v.getBoolean("bl"));//粗体显示
          // 设置字体斜体 0 常规 、 1 斜体
          font.setItalic(cellObject_v.getBoolean("it") != null && cellObject_v.getBoolean("it"));//斜体
          // 设置字体样式.默认0   0 Times New Roman、 1 Arial、2 Tahoma 、3 Verdana、4 微软雅黑、5 宋体（Song）、6 黑体（ST Heiti）、
          // 7 楷体（ST Kaiti）、8 仿宋（ST FangSong）、9 新宋体（ST Song）、10 华文新魏、11 华文行楷、12 华文隶书
          font.setFontName(FontMap.get(cellObject_v.getInteger("ff") == null ? 0 : cellObject_v.getInteger("ff")));
          // 设置字体大小 默认14
          font.setFontHeightInPoints((short) (cellObject_v.getInteger("fs") == null ? 14 : cellObject_v.getInteger("fs")));
          // 字体颜色
          String fc = cellObject_v.getString("fc") == null ? "" : cellObject_v.getString("fc");
          if (fc.length() > 0) {
            font.setColor(new XSSFColor(setColor(fc), new DefaultIndexedColorMap()));
          }
          // 设置删除线
          font.setStrikeout(cellObject_v.getInteger("cl") != null && cellObject_v.getBoolean("cl"));

          style.setFont(font);

          //*********************** 设置对齐方式 **********************
          //垂直对齐    0 中间、1 上、2下
          int vt = cellObject_v.getInteger("vt") == null ? 1 : cellObject_v.getInteger("vt");
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
          // 水平对齐 0 居中、1 左、2右
          int ht = cellObject_v.getInteger("ht") == null ? 1 : cellObject_v.getInteger("ht");
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

          //设置自动换行
          // 文本换行 0 截断、1溢出、2 自动换行
          int tb = cellObject_v.getInteger("tb") == null ? 0 : cellObject_v.getInteger("tb");
          switch (tb) {
            case 2:
              style.setWrapText(true);
              break;
            case 0:
            case 1:
              // 暂不处理
              break;
          }

          // 竖排文字
          int tr = cellObject_v.getInteger("tr") == null ? 0 : cellObject_v.getInteger("tr");
          switch (tr) {
            case 1:
              style.setRotation((short) 45);
              break;
            case 2:
              style.setRotation((short) -45);
              break;
            case 3:
              style.setRotation((short) 180);
              break;
            case 4:
              style.setRotation((short) 90);
              break;
            case 5:
              style.setRotation((short) -90);
              break;
          }

          // 批注
          JSONObject ps = cellObject_v.getJSONObject("ps") == null ? null : cellObject_v.getJSONObject("ps");


          // 将样式配置到单元格
          cell.setCellStyle(style);


        }
      } else {
        String str_ = cellObject.get("r") + "_" + cellObject.get("c") + "=" + ((JSONObject) cellObject.get("v")).get("v") + "\n";
        System.out.println("错误的=" + index + ">>>" + str_);
      }
    }
    //设置边框
    XSSFSheet this_sheet = setBorder(borderInfoObjectList, sheet);

    return this_sheet;
  }

  //设置边框
  private static XSSFSheet setBorder(JSONArray borderInfoObjectList, XSSFSheet sheet) {

    //一定要通过 cell.getCellStyle()  不然的话之前设置的样式会丢失
    //设置边框
    if (null != borderInfoObjectList) {
      for (int i = 0; i < borderInfoObjectList.size(); i++) {
        JSONObject borderInfoObject = (JSONObject) borderInfoObjectList.get(i);
        //单个单元格
        if (borderInfoObject.get("rangeType").equals("cell")) {
          JSONObject borderValueObject = borderInfoObject.getJSONObject("value");

          JSONObject l = borderValueObject.getJSONObject("l");// 左边框
          JSONObject r = borderValueObject.getJSONObject("r");// 右边框
          JSONObject t = borderValueObject.getJSONObject("t");// 上边框
          JSONObject b = borderValueObject.getJSONObject("b");// 下边框


          int row = borderValueObject.getInteger("row_index");
          int col = borderValueObject.getInteger("col_index");

          XSSFCell cell = sheet.getRow(row).getCell(col);

          if (l != null) {  //左边框
            cell.getCellStyle().setBorderLeft(BordMap.get(l.getInteger("style"))); // 样式
            cell.getCellStyle().setLeftBorderColor(new XSSFColor(setColor(l.getString("color")), DEFAULT_INDEXED_COLOR_MAP));  //颜色
          }
          if (r != null) {  //右边框
            cell.getCellStyle().setBorderRight(BordMap.get(r.getInteger("style"))); // 样式
            cell.getCellStyle().setRightBorderColor(new XSSFColor(setColor(r.getString("color")), DEFAULT_INDEXED_COLOR_MAP));  //颜色
          }
          if (t != null) {  //顶部边框
            cell.getCellStyle().setBorderTop(BordMap.get(t.getInteger("style"))); // 样式
            cell.getCellStyle().setTopBorderColor(new XSSFColor(setColor(t.getString("color")), DEFAULT_INDEXED_COLOR_MAP));  //颜色
          }
          if (b != null) {  //底部边框
            cell.getCellStyle().setBorderBottom(BordMap.get(b.getInteger("style"))); // 样式
            cell.getCellStyle().setBottomBorderColor(new XSSFColor(setColor(b.getString("color")), DEFAULT_INDEXED_COLOR_MAP));  //颜色
          }
        } else if (borderInfoObject.get("rangeType").equals("range")) {//选区
          /***
           String color_Hex = borderInfoObject.getString("color");
           // TODO color_hex转换成int,暂未找到方法，android.graph.color的parseColor可能；
           int color = 1;
           int style_ = borderInfoObject.getInteger("style");
           String borderType = borderInfoObject.getString("borderType");  // TODO 边框类型没有设置，需要完善

           JSONObject rangObject = (JSONObject) ((JSONArray) (borderInfoObject.get("range"))).get(0);
           JSONArray rowList = rangObject.getJSONArray("row");
           JSONArray columnList = rangObject.getJSONArray("column");

           CellRangeAddress borderRegion = new CellRangeAddress(rowList.getInteger(0), rowList.getInteger(1),
           columnList.getInteger(0), columnList.getInteger(1));
           if (borderType.equalsIgnoreCase("border-left")) {

           RegionUtil.setBorderLeft(BordMap.get(style_), borderRegion, sheet);
           RegionUtil.setLeftBorderColor(color, borderRegion, sheet);
           } else if (borderType.equalsIgnoreCase("border-right")) {

           RegionUtil.setBorderRight(BordMap.get(style_), borderRegion, sheet);
           RegionUtil.setRightBorderColor(color, borderRegion, sheet);
           } else if (borderType.equalsIgnoreCase("border-top")) {

           RegionUtil.setBorderTop(BordMap.get(style_), borderRegion, sheet);
           RegionUtil.setTopBorderColor(color, borderRegion, sheet);
           } else if (borderType.equalsIgnoreCase("border-bottom")) {

           RegionUtil.setBorderBottom(BordMap.get(style_), borderRegion, sheet);
           RegionUtil.setBottomBorderColor(color, borderRegion, sheet);
           } else if (borderType.equalsIgnoreCase("border-all") || borderType.equalsIgnoreCase("border-outside")) {
           RegionUtil.setBorderBottom(BordMap.get(style_), borderRegion, sheet);
           RegionUtil.setBorderTop(BordMap.get(style_), borderRegion, sheet);
           RegionUtil.setBorderRight(BordMap.get(style_), borderRegion, sheet);
           RegionUtil.setBorderLeft(BordMap.get(style_), borderRegion, sheet);

           RegionUtil.setBottomBorderColor(color, borderRegion, sheet);
           RegionUtil.setTopBorderColor(color, borderRegion, sheet);
           RegionUtil.setRightBorderColor(color, borderRegion, sheet);
           RegionUtil.setLeftBorderColor(color, borderRegion, sheet);
           } else if (borderType.equalsIgnoreCase("border-horizontal")) {
           RegionUtil.setBorderBottom(BordMap.get(style_), borderRegion, sheet);
           RegionUtil.setBorderTop(BordMap.get(style_), borderRegion, sheet);

           RegionUtil.setBottomBorderColor(color, borderRegion, sheet);
           RegionUtil.setTopBorderColor(color, borderRegion, sheet);
           } else if (borderType.equalsIgnoreCase("border-vertical")) {
           RegionUtil.setBorderRight(BordMap.get(style_), borderRegion, sheet);
           RegionUtil.setBorderLeft(BordMap.get(style_), borderRegion, sheet);

           RegionUtil.setRightBorderColor(color, borderRegion, sheet);
           RegionUtil.setLeftBorderColor(color, borderRegion, sheet);
           } else if (borderType.equalsIgnoreCase("border-inside") || borderType.equalsIgnoreCase("border-none")) {
           // 暂不处理
           }
           */
        }
      }
    }

    return sheet;
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


  // 颜色转换
  private static Color setColor(String color) {
    if (color.startsWith("#")) {
      int[] hexColor = hex2RGB(color);
      return new Color(hexColor[0], hexColor[1], hexColor[2]);
    } else if (color.startsWith("rgb")) {
      // rgb颜色转换
      String[] rgb_str = color.replace("rgb(", "").replace(")", "").replace(" ", "").split(",");
      return new Color(Integer.parseInt(rgb_str[0]), Integer.parseInt(rgb_str[1]), Integer.parseInt(rgb_str[2]));
    } else {
      throw new ValueException("{}--this color format is not supported now", color);
    }
  }

  /**
   * 16进制颜色字符串转换成rgb
   *
   * @param hexStr
   * @return rgb
   */
  public static int[] hex2RGB(String hexStr) {
    if (hexStr != null && !"".equals(hexStr) && hexStr.length() == 7) {
      int[] rgb = new int[3];
      rgb[0] = Integer.valueOf(hexStr.substring(1, 3), 16);
      rgb[1] = Integer.valueOf(hexStr.substring(3, 5), 16);
      rgb[2] = Integer.valueOf(hexStr.substring(5, 7), 16);
      return rgb;
    }
    return null;
  }


}

// TODO：
//  1、char;
//  2、单元格注释；
//  3、单元格复选框
//  4、透视表；
//  5、公式；
//  6、表的冻结；
//  7、表的筛选；
//  8、删除线；
//  9、边框设置；
//  10、字体旋转倾斜
//  11、
