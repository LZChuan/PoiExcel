package com.chuan;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import jdk.nashorn.internal.runtime.regexp.joni.exception.ValueException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import com.chuan.chartsUtils.*;

import java.awt.Color;
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

  public static XSSFWorkbook exportLuckySheetByPOI(String excelData) {
    //去除luckysheet中 &#xA 的换行
    excelData = excelData.replace("&#xA;", "\\r\\n");
    JSONArray jsonArray = JSONObject.parseArray(excelData);
    //创建操作Excel的XSSFWorkbook对象
    XSSFWorkbook excel = new XSSFWorkbook();

    //有多少个表就循环多少次
    for (int sheetIndex = 0; sheetIndex < jsonArray.size(); sheetIndex++) {
      //获取配置
      JSONObject jsonObject = jsonArray.getJSONObject(sheetIndex);
      //创建XSSFSheet对象并命名
      XSSFSheet sheet = excel.createSheet(jsonObject.getString("name"));

      // 过滤掉数据透视sheet
      boolean isPivotTable = jsonObject.getBoolean("isPivotTable") != null && jsonObject.getBoolean("isPivotTable");
      if (isPivotTable) {
        continue;
      }

      // 给sheet填充数据
      createRowsAndColumns(excel, sheet, jsonObject);
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
      // 图表数据
      JSONArray chartData = jsonObject.getJSONArray("chart");
      if (chartData != null && chartData.size() > 0) {
        DrawSheetCharts.SheetCharts(sheet, chartData);
      }
    }
    return excel;
  }

  private static void createRowsAndColumns(
    XSSFWorkbook excel, XSSFSheet sheet, JSONObject sheetObject) {

    //更新和存储使用的单元格数据
    JSONArray data = sheetObject.getJSONArray("data");
    // 表的整体配置
    JSONObject config = sheetObject.getJSONObject("config");

    //行高
    JSONObject rowlen = config.getJSONObject("rowlen");
    //列宽
    JSONObject columnlen = config.getJSONObject("columnlen");
    //边框信息
    JSONArray borderInfo = config.getJSONArray("borderInfo");
    //隐藏行
    JSONObject rowHidden = config.getJSONObject("rowhidden");

    //  设置列宽
    for (int j = 0; j < data.getJSONArray(0).size(); j++) {
      // 设置列宽
      if (columnlen != null && columnlen.getInteger(String.valueOf(j)) != null) {
        sheet.setColumnWidth(j, columnlen.getInteger(j + "") * 42);//列宽px值
      } else {
        sheet.setColumnWidth(j, 75 * 42);//默认列宽px值
      }
    }

    // 创建 行
    XSSFRow row;
    for (int i = 0; i < data.size(); i++) {
      // 创建行并设置行高
      row = sheet.createRow(i);
      // ######################### 设置行高,默认为20  ################
      float rowhigh = 20.0f;
      if (rowlen != null && rowlen.getInteger(String.valueOf(i)) != null) {
        rowhigh = rowlen.getFloat(String.valueOf(i));
      }
      if (rowHidden != null && rowHidden.getInteger(String.valueOf(i)) != null) {
        rowhigh = 0;
      }
      row.setHeightInPoints(rowhigh);

      //遍历列
      for (int j = 0; j < data.getJSONArray(i).size(); j++) {
        // 创建单元格
        row.createCell(j);
      }
    }
    //设置所有单元格值
    setCellValue(sheetObject, borderInfo, sheet, excel);
  }

  private static void setCellValue(JSONObject sheetObject, JSONArray borderInfoObjectList, XSSFSheet sheet, XSSFWorkbook excel) {
    //初始化的数据
    JSONArray cellAllData = sheetObject.getJSONArray("data");
    JSONObject dataVerification = sheetObject.getJSONObject("dataVerification");

    String cellType = "";
    String cellFormat;

    // 设置所有单元格信息
    for (int row_idx = 0; row_idx < cellAllData.size(); row_idx++) {
      // 行数据
      JSONArray cellRowData = cellAllData.getJSONArray(row_idx);
      for (int col_idx = 0; col_idx < cellRowData.size(); col_idx++) {
        // 单元格数据
        JSONObject cellRowColData = cellRowData.getJSONObject(col_idx);
        if (cellRowColData == null) {
          continue;
        }
        // 样式
        XSSFCellStyle style = excel.createCellStyle();
        // 字体
        XSSFFont font = excel.createFont();//字体样式
        // 内容格式
        XSSFDataFormat dataFormat = excel.createDataFormat();

        XSSFCell cell = sheet.getRow(row_idx).getCell(col_idx);

        //************************* 字体设置 ************************
        // 设置字体是否加粗 	0 常规, 1加粗
        font.setBold(cellRowColData.getBoolean("bl") != null && cellRowColData.getBoolean("bl"));//粗体显示
        // 设置字体斜体 0 常规 、 1 斜体
        font.setItalic(cellRowColData.getBoolean("it") != null && cellRowColData.getBoolean("it"));//斜体
        // 设置字体样式.默认0   0 Times New Roman、 1 Arial、2 Tahoma 、3 Verdana、4 微软雅黑、5 宋体（Song）、6 黑体（ST Heiti）、
        // 7 楷体（ST Kaiti）、8 仿宋（ST FangSong）、9 新宋体（ST Song）、10 华文新魏、11 华文行楷、12 华文隶书
        String fontStr;
        if (cellRowColData.getString("ff") != null) {
          try {
            Integer fontNum = cellRowColData.getInteger("ff");
            fontStr = FontMap.get(fontNum);
          } catch (Exception e) {
            fontStr = cellRowColData.getString("ff");
          }
        } else {
          fontStr = FontMap.get(0);
        }
        font.setFontName(fontStr);
        // 设置字体大小 默认14
        font.setFontHeightInPoints((short) (cellRowColData.getInteger("fs") == null ? 14 : cellRowColData.getInteger("fs")));
        // 字体颜色
        String fc = cellRowColData.getString("fc") == null ? "" : cellRowColData.getString("fc");
        if (fc.length() > 0) {
          font.setColor(new XSSFColor(setColor(fc), new DefaultIndexedColorMap()));
        }
        // 设置删除线
        font.setStrikeout(cellRowColData.getInteger("cl") != null && cellRowColData.getBoolean("cl"));
        style.setFont(font);

        //单元格内容 类型
        if (cellRowColData.containsKey("ct")) {
          // Type类型
          cellType = cellRowColData.getJSONObject("ct").getString("t");
          // Format格式的定义串
          cellFormat = cellRowColData.getJSONObject("ct").getString("fa");
          if (cellFormat != null && !Objects.equals(cellFormat, "")) {
            style.setDataFormat(dataFormat.getFormat(cellFormat));
          }
          // 单元格数据验证设置
          if (dataVerification != null && dataVerification.size() > 0) {
            DataVerificationFun.setCellVerfication(excel, sheet, dataVerification, cellFormat);
          }
        }

        // 单元格内容填充
        String cellValue;
        //如果有公式，则使用公式
        if (cellRowColData.get("f") != null) {
          try {
            cellValue = cellRowColData.getString("f");
            cell.setCellFormula(cellValue.substring(1));
          } catch (Exception e) {
            e.printStackTrace();
            System.out.println("this formula can not be process");
          }
        }
        // 设置cell中的值
        if (cellRowColData.get("v") != null) {
          cellValue = cellRowColData.getString("v");
          // 根据单元格 类型 修改单元格内容的值
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
        if (cellRowColData.containsKey("mc")) {
          JSONObject mergeObject = (JSONObject) cellRowColData.get("mc");
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
        if (cellRowColData.getString("bg") != null) {
          style.setFillPattern(FillPatternType.SOLID_FOREGROUND);    //设置填充方案
          style.setFillForegroundColor(new XSSFColor(setColor(cellRowColData.getString("bg")), new DefaultIndexedColorMap()));
        }

        //*********************** 设置对齐方式 **********************
        //垂直对齐    0 中间、1 上、2下
        int vt = cellRowColData.getInteger("vt") == null ? 1 : cellRowColData.getInteger("vt");
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
        int ht = cellRowColData.getInteger("ht") == null ? 1 : cellRowColData.getInteger("ht");
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
        int tb = cellRowColData.getInteger("tb") == null ? 0 : cellRowColData.getInteger("tb");
        if (tb == 2) {
          style.setWrapText(true);
        }

        // 竖排文字
        int tr = cellRowColData.getInteger("tr") == null ? 0 : cellRowColData.getInteger("tr");
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
        if (cellRowColData.containsKey("ps")) {
          JSONObject ps = cellRowColData.getJSONObject("ps");
          XSSFDrawing commenting = sheet.createDrawingPatriarch();
          XSSFComment cellComment = commenting.createCellComment(commenting.createAnchor(0, 0, 0, 0,
            0, 0, 0, 0));
          cellComment.setString(new XSSFRichTextString(ps.getString("value")));
          cell.setCellComment(cellComment);
        }

        // 将样式配置到单元格
        cell.setCellStyle(style);
      }
    }
    //设置边框
    setBorder(borderInfoObjectList, sheet);
  }

  //设置边框
  private static void setBorder(JSONArray borderInfoObjectList, XSSFSheet sheet) {

    //设置边框
    if (null != borderInfoObjectList) {
      for (Object o : borderInfoObjectList) {
        JSONObject borderInfoObject = (JSONObject) o;
        //单个单元格
        if (borderInfoObject.get("rangeType").equals("cell")) {
          JSONObject borderValueObject = borderInfoObject.getJSONObject("value");

          JSONObject l = borderValueObject.getJSONObject("l");// 左边框
          JSONObject r = borderValueObject.getJSONObject("r");// 右边框
          JSONObject t = borderValueObject.getJSONObject("t");// 上边框
          JSONObject b = borderValueObject.getJSONObject("b");// 下边框

          //一定要通过 cell.getCellStyle()  不然的话之前设置的样式会丢失
          XSSFCell cell = sheet.getRow(borderValueObject.getInteger("row_index"))
            .getCell(borderValueObject.getInteger("col_index"));
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
        }
      }
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
