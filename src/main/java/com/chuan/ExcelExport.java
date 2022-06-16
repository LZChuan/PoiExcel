package com.chuan;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import jdk.nashorn.internal.runtime.regexp.joni.exception.ValueException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
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
    FontMap.put(-1, "Calibri");
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
      JSONArray chartData = jsonObject.getJSONArray("chart");
      // 表的整体配置
      JSONObject config = jsonObject.getJSONObject("config");

      boolean isPivotTable = jsonObject.getBoolean("isPivotTable") != null && jsonObject.getBoolean("isPivotTable");
      if (isPivotTable) {   // TODO 透视表暂时不支持
        continue;
      }
      //单元格的样式
      //      XSSFCellStyle cellStyle = excel.createCellStyle();
      //创建XSSFSheet对象并命名
      XSSFSheet sheet = excel.createSheet(jsonObject.getString("name"));
      // 给sheet填充数据
      sheet = createRowsAndColumns(excel, sheet, data, config, celldata);

      // 给 sheet绘图
      if (chartData != null && chartData.size() > 0) {
        // 目前仅折线图
        sheet = DrawSheetCharts.SheetCharts(sheet, chartData);

//        for (int chartIndex = 0; chartIndex < chartData.size(); chartIndex++) {
//          JSONObject thisChartData = chartData.getJSONObject(chartIndex);
//          JSONObject chartOptions = thisChartData.getJSONObject("chartOptions");
//          JSONObject chartDefaultOption = chartOptions.getJSONObject("defaultOption");
//
//          if (!chartOptions.getString("chartAllType").split("\\|")[1].equalsIgnoreCase("line")) {
//            continue;
//          }
//
//          // 像素位置转换成 行列位置
//          List<Integer> anchor_int = Utils.Pix2Anchor(thisChartData.getInteger("width"), thisChartData.getInteger("height"), thisChartData.getInteger("left"), thisChartData.getInteger("top"));
//          // 创建一个画布
//          XSSFDrawing drawing = sheet.createDrawingPatriarch();
//          // 创建一个chart对象
//          XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0, anchor_int.get(0), anchor_int.get(1), anchor_int.get(2), anchor_int.get(3)));
//
//          // **********************************  标题 轴 图例 配置 ***********************
//          // 对defaultOption进行解析与设置
//
//          // 标题
//          if (chartDefaultOption.getJSONObject("title").getBoolean("show")) {
//            // 获取title配置
//            JSONObject chartTitle = chartDefaultOption.getJSONObject("title");
//            // title内容
//            chart.setTitleText(chartTitle.getString("text"));
//            //        chart.setTitleOverlay(true);        // 标题覆盖
//            // TODO 待完善
//          }
//          // 子标题
//          if (chartDefaultOption.getJSONObject("subtitle").getBoolean("show")) {
//            // 获取子标题配置
//            JSONObject chartSubtitle = chartDefaultOption.getJSONObject("subtitle");
//            // TODO 待完善
//          }
//          // 图例
//          if (chartDefaultOption.getJSONObject("legend").getBoolean("show")) {
//            // 获取图例配置
//            JSONObject chartLegend = chartDefaultOption.getJSONObject("legend");
//            // 图例位置
//            XDDFChartLegend legend = chart.getOrAddLegend();
//            String position = chartLegend.getJSONObject("position").getString("value");
//            if (position.equalsIgnoreCase("right")) {
//              legend.setPosition(LegendPosition.RIGHT);
//            } else if (position.equalsIgnoreCase("left")) {
//              legend.setPosition(LegendPosition.LEFT);
//            } else if (position.equalsIgnoreCase("BOTTOM")) {
//              legend.setPosition(LegendPosition.BOTTOM);
//            } else {
//              legend.setPosition(LegendPosition.TOP);
//            }
//            // TODO 待完善
//          }
//          // 坐标轴
//          XDDFCategoryAxis xAxisPosition = null;    // x轴
//          XDDFCategoryDataSource xAxisValue = null;
//          XDDFValueAxis yAxisPosition = null;       // y轴
//          XDDFCategoryDataSource yAxisValue = null;
//          if (chartDefaultOption.getJSONObject("axis") != null) {
//            JSONObject xAxisUp = chartDefaultOption.getJSONObject("axis").getJSONObject("xAxisUp");
//            JSONObject xAxisDown = chartDefaultOption.getJSONObject("axis").getJSONObject("xAxisDown");
//            JSONObject yAxisLeft = chartDefaultOption.getJSONObject("axis").getJSONObject("yAxisLeft");
//            JSONObject yAxisRight = chartDefaultOption.getJSONObject("axis").getJSONObject("yAxisRight");
//            // x轴
//            if (xAxisUp.getBoolean("show")) {
//              xAxisPosition = chart.createCategoryAxis(AxisPosition.TOP);
//              if (xAxisUp.getJSONObject("title").getBoolean("showTitle")) {
//                xAxisPosition.setTitle(xAxisUp.getJSONObject("title").getString("text"));
//              }
//              if (xAxisUp.containsKey("data")) {
//                xAxisValue = XDDFDataSourcesFactory.fromArray(Utils.JsonArray2ArrayString(xAxisUp.getJSONArray("data")));
//              }
//            }
//            if (xAxisDown.getBoolean("show")) {
//              xAxisPosition = chart.createCategoryAxis(AxisPosition.BOTTOM);
//              if (xAxisDown.getJSONObject("title").getBoolean("showTitle")) {
//                xAxisPosition.setTitle(xAxisDown.getJSONObject("title").getString("text"));
//              }
//              if (xAxisDown.containsKey("data")) {
//                xAxisValue = XDDFDataSourcesFactory.fromArray(Utils.JsonArray2ArrayString(xAxisDown.getJSONArray("data")));
//              }
//            }
//            // y轴
//            if (yAxisLeft.getBoolean("show")) {
//              yAxisPosition = chart.createValueAxis(AxisPosition.LEFT);
//              if (yAxisLeft.getJSONObject("title").getBoolean("showTitle")) {
//                yAxisPosition.setTitle(yAxisLeft.getJSONObject("title").getString("text"));
//              }
//              if (yAxisLeft.containsKey("data")) {
//                yAxisValue = XDDFDataSourcesFactory.fromArray(Utils.JsonArray2ArrayString(yAxisLeft.getJSONArray("data")));
//              }
//            }
//            if (yAxisRight.getBoolean("show")) {
//              yAxisPosition = chart.createValueAxis(AxisPosition.RIGHT);
//              if (yAxisRight.getJSONObject("title").getBoolean("showTitle")) {
//                yAxisPosition.setTitle(yAxisRight.getJSONObject("title").getString("text"));
//              }
//              if (yAxisRight.containsKey("data")) {
//                yAxisValue = XDDFDataSourcesFactory.fromArray(Utils.JsonArray2ArrayString(yAxisRight.getJSONArray("data")));
//              }
//            }
//          }
//
//          // 绘图
//          // LINE：折线图，
//
//          XDDFLineChartData draw_data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, xAxisPosition, yAxisPosition);
//
//          // 添加各个折线
//          if (chartDefaultOption.getJSONArray("series") != null) {
//            JSONArray series = chartDefaultOption.getJSONArray("series");
//            for (int series_idx = 0; series_idx < series.size(); series_idx++) {
//              JSONObject series_line = series.getJSONObject(series_idx);
//
//              // 1条折线的数据
//              XDDFNumericalDataSource<Integer> plot_data = XDDFDataSourcesFactory.fromArray(Utils.JsonArray2ArrayDouble(series_line.getJSONArray("data")));
//              // 生成 1条折线
//              XDDFLineChartData.Series series_plot = (XDDFLineChartData.Series) draw_data.addSeries(xAxisValue, plot_data);
//              // 折线的标题
//              series_plot.setTitle(series_line.getString("name"), null);
//              if (series_line.getString("type") != null) {
//                if (series_line.getString("type").equalsIgnoreCase("line")) {
//                  series_plot.setSmooth(false);              // 折线样式---直线
//                }
//              }
//              // mark 暂无
//              //          series_plot.setMarkerSize((short) 6);          // 设置标记大小
//              //          series_plot.setMarkerStyle(MarkerStyle.STAR);  // 设置标记样式，星星
//
//            }
//          }
//
//          // 绘制
//          chart.plot(draw_data);
//        }

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
      // 创建行并设置行高
      row = sheet.createRow(i);
      if (rowlen != null) {        //设置行高,默认为20
        row.setHeightInPoints(rowlen.getInteger(String.valueOf(i)) == null ? 20 : rowlen.getInteger(String.valueOf(i)));
      } else {
        row.setHeightInPoints(20f);
      }
      //创建列
      for (int j = 0; j < data.getJSONArray(i).size(); j++) {
        /*********/
        // 设置列宽，无默认值  TODO 能否在外层设置？
        if (columnlen != null && columnlen.getInteger(String.valueOf(j)) != null) {
          sheet.setColumnWidth(j, columnlen.getInteger(j + "") * 42);//列宽px值
        }
        //这里可以设置celltype，在构造方法里
        /********/
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

      XSSFCell cell = null;

      JSONObject cellObject = cellData.getJSONObject(index);


      JSONObject cellObject_v = null;
      try {

        cellObject_v = cellObject.getJSONObject("v");
      } catch (Exception e) {
        e.printStackTrace();
      }


      if (sheet.getRow((int) cellObject.get("r")) != null && sheet.getRow((int) cellObject.get("r")).getCell((int) cellObject.get("c")) != null) {

        cell = sheet.getRow((int) cellObject.get("r")).getCell((int) cellObject.get("c"));

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

        // 获取单元格 内容  TODO 需要完善 公式、value值、m值的取舍；
        String value = "";
        if (cellObject_v.get("v") != null) {
          value = cellObject_v.getString("v");
        }
        //如果有公式，设置公式
        if (cellObject_v.get("f") != null) {
          // TODO 有很多公式无法支持，使用try进行设置
          String value_t = value;
          try {
            value = cellObject_v.getString("f");
            cell.setCellFormula(value.substring(1));//不需要=符号  TODO 公式 需要验证
          } catch (Exception e) {
            e.printStackTrace();
            value = value_t;
          }
        }
        // 根据单元格 类型 修改单元格内容的值  TODO 需要完善，识别 各种类型；
        if (!value.equals("")) {
          switch (cellType) {
            case "s": // 纯文本
            case "g": // 默认格式
              cell.setCellValue(value);
              break;

            case "d": // 时间
              try {
                cell.setCellValue(Double.parseDouble(value));
              } catch (Exception e) {
                e.printStackTrace();
              }
              break;

            case "n": // 数字 货币， TODO 遇到计算公式怎么处理？
              try {
                cell.setCellValue(Double.parseDouble(value));
              } catch (Exception e) {
                e.printStackTrace();
                // 会出现 公式的情况
                cell.setCellValue(value);
              }
              break;
          }
        }

        // 分割单元格 TODO 暂缺


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

        //设置合并单元格的样式有问题
        int ff = cellObject_v.getInteger("ff") == null ? -1 : cellObject_v.getInteger("ff");//0 Times New Roman、 1 Arial、2 Tahoma 、3 Verdana、4 微软雅黑、5 宋体（Song）、6 黑体（ST Heiti）、7 楷体（ST Kaiti）、 8 仿宋（ST FangSong）、9 新宋体（ST Song）、10 华文新魏、11 华文行楷、12 华文隶书
        int fs = cellObject_v.getInteger("fs") == null ? 14 : cellObject_v.getInteger("fs");//字体大小
        int bl = cellObject_v.getInteger("bl") == null ? 0 : cellObject_v.getInteger("bl");//粗体	0 常规 、 1加粗
        int it = cellObject_v.getInteger("it") == null ? 0 : cellObject_v.getInteger("it");//斜体	0 常规 、 1 斜体
        String fc = cellObject_v.getString("fc") == null ? "" : cellObject_v.getString("fc");//字体颜色
        int vt = cellObject_v.getInteger("vt") == null ? 1 : cellObject_v.getInteger("vt");//垂直对齐	 0 中间、1 上、2下
        int ht = cellObject_v.getInteger("ht") == null ? 1 : cellObject_v.getInteger("ht");//0 居中、1 左、2右

        //************************* 字体设置 ************************
        // 设置字体是否加粗
        font.setBold(bl == 1);//粗体显示
        // 设置字体斜体
        font.setItalic(it == 1);//斜体
        // 设置字体样式，默认为Calibri
        font.setFontName(FontMap.get(ff));//字体名字
        // 设置字体大小
        font.setFontHeightInPoints((short) fs);//字体大小
        // 字体颜色
        if (fc.length() > 0) {
          font.setColor(new XSSFColor(setColor(fc), new DefaultIndexedColorMap()));
        }
        style.setFont(font);

        //*********************** 设置对齐方式 **********************
        //设置垂直水平对齐方式
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

        //设置自动换行
        style.setWrapText(true);
        // 将样式配置到单元格
        cell.setCellStyle(style);

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

          JSONObject l = borderValueObject.getJSONObject("l");
          JSONObject r = borderValueObject.getJSONObject("r");
          JSONObject t = borderValueObject.getJSONObject("t");
          JSONObject b = borderValueObject.getJSONObject("b");


          int row = borderValueObject.getInteger("row_index");
          int col = borderValueObject.getInteger("col_index");

          XSSFCell cell = sheet.getRow(row).getCell(col);


          if (l != null) {  //左边框
            cell.getCellStyle().setBorderLeft(BordMap.get(l.getInteger("style"))); // 样式
            cell.getCellStyle().setLeftBorderColor(new XSSFColor(setColor(l.getString("color")), new DefaultIndexedColorMap()));  //颜色
          }
          if (r != null) {  //右边框
            cell.getCellStyle().setBorderRight(BordMap.get(r.getInteger("style"))); // 样式
            cell.getCellStyle().setRightBorderColor(new XSSFColor(setColor(r.getString("color")), new DefaultIndexedColorMap()));  //颜色
          }
          if (t != null) {  //顶部边框
            cell.getCellStyle().setBorderTop(BordMap.get(t.getInteger("style"))); // 样式
            cell.getCellStyle().setTopBorderColor(new XSSFColor(setColor(t.getString("color")), new DefaultIndexedColorMap()));  //颜色
          }
          if (b != null) {  //底部边框
            cell.getCellStyle().setBorderBottom(BordMap.get(b.getInteger("style"))); // 样式
            cell.getCellStyle().setBottomBorderColor(new XSSFColor(setColor(b.getString("color")), new DefaultIndexedColorMap()));  //颜色
          }
        } else if (borderInfoObject.get("rangeType").equals("range")) {//选区
          String bg_ = borderInfoObject.getString("color");
          int style_ = borderInfoObject.getInteger("style");

          String borderType = borderInfoObject.getString("borderType");  // TODO 边框类型没有设置，需要完善

          JSONObject rangObject = (JSONObject) ((JSONArray) (borderInfoObject.get("range"))).get(0);
          JSONArray rowList = rangObject.getJSONArray("row");
          JSONArray columnList = rangObject.getJSONArray("column");

          for (int row_ = rowList.getInteger(0); row_ < rowList.getInteger(rowList.size() - 1) + 1; row_++) {
            for (int col_ = columnList.getInteger(0); col_ < columnList.getInteger(columnList.size() - 1) + 1; col_++) {
              XSSFCell cell = sheet.getRow(row_).getCell(col_);

              XSSFColor color_tmp = new XSSFColor(setColor(bg_), DEFAULT_INDEXED_COLOR_MAP);
              BorderStyle style_tmp = BordMap.get(style_);

              cell.getCellStyle().setBorderLeft(style_tmp); //左边框
              cell.getCellStyle().setBorderRight(style_tmp); //右边框
              cell.getCellStyle().setBorderTop(style_tmp); //顶部边框
              cell.getCellStyle().setBorderBottom(style_tmp); //底部边框

              cell.getCellStyle().setLeftBorderColor(color_tmp);//左边框颜色
              cell.getCellStyle().setRightBorderColor(color_tmp);//右边框颜色
              cell.getCellStyle().setTopBorderColor(color_tmp);//顶部边框颜色
              cell.getCellStyle().setBottomBorderColor(color_tmp);//底部边框颜色
            }
          }
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
      int hexColor = Integer.parseInt(color.replace("#", ""), 16);
      return new Color(hexColor);
    } else if (color.startsWith("rgb")) {
      // rgb颜色转换
      String[] rgb_str = color.replace("rgb(", "").replace(")", "").replace(" ", "").split(",");
      return new Color(Integer.parseInt(rgb_str[0]), Integer.parseInt(rgb_str[1]), Integer.parseInt(rgb_str[2]));
    } else {
      throw new ValueException("{}--this color format is not supported now", color);
    }
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
