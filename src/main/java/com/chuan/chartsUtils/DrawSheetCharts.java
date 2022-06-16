package com.chuan.chartsUtils;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.util.List;

public class DrawSheetCharts {

  public static XSSFSheet SheetCharts(XSSFSheet sheet, JSONArray chartData) {

    for (int chartIndex = 0; chartIndex < chartData.size(); chartIndex++) {
      JSONObject thisChartData = chartData.getJSONObject(chartIndex);
      String chartType = thisChartData.getJSONObject("chartOptions").getString("chartAllType").split("\\|")[1];
      switch (chartType.toLowerCase()) {
        case "line":
          sheet = getLineChart(sheet, thisChartData);
          break;
        case "pie":
          sheet = getPieChart(sheet,thisChartData);
          break;
        case "column":
//          sheet = getColumnChart(sheet, thisChartData);
          break;
      }
    }
    return sheet;
  }


  private static XSSFSheet getLineChart(XSSFSheet sheet, JSONObject chartData) {

    JSONObject chartOptions = chartData.getJSONObject("chartOptions");
    JSONObject chartDefaultOption = chartOptions.getJSONObject("defaultOption");

    // 创建一个画布
    XSSFDrawing drawing = sheet.createDrawingPatriarch();

    // 像素位置转换成 行列位置
    List<Integer> anchor_int = Utils.Pix2Anchor(
        chartData.getInteger("width"),
        chartData.getInteger("height"),
        chartData.getInteger("left"),
        chartData.getInteger("top"));
    // 创建一个chart对象
    XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0,
        anchor_int.get(0), anchor_int.get(1), anchor_int.get(2), anchor_int.get(3)));

    // **********************************  标题 轴 图例 配置 ***********************
    // 对defaultOption进行解析与设置

    // 标题
    if (chartDefaultOption.getJSONObject("title").getBoolean("show")) {
      // 获取title配置
      JSONObject chartTitle = chartDefaultOption.getJSONObject("title");
      // title内容
      chart.setTitleText(chartTitle.getString("text"));
      //        chart.setTitleOverlay(true);        // 标题覆盖
      // TODO 待完善
    }

    // 子标题  TODO 暂未找到如何设置
    if (chartDefaultOption.getJSONObject("subtitle").getBoolean("show")) {
      // 获取子标题配置
      JSONObject chartSubtitle = chartDefaultOption.getJSONObject("subtitle");
      // TODO 待完善
    }

    // 图例
    if (chartDefaultOption.getJSONObject("legend").getBoolean("show")) {
      // 获取图例配置
      JSONObject chartLegend = chartDefaultOption.getJSONObject("legend");
      // 图例位置
      XDDFChartLegend legend = chart.getOrAddLegend();
      String position = chartLegend.getJSONObject("position").getString("value");
      if (position.equalsIgnoreCase("right")) {
        legend.setPosition(LegendPosition.RIGHT);
      } else if (position.equalsIgnoreCase("left")) {
        legend.setPosition(LegendPosition.LEFT);
      } else if (position.equalsIgnoreCase("BOTTOM")) {
        legend.setPosition(LegendPosition.BOTTOM);
      } else {
        legend.setPosition(LegendPosition.TOP);
      }
      // TODO 待完善
    }

    // 坐标轴   TODO 样式设置
    XDDFCategoryAxis xAxisPosition = null;    // x轴
    XDDFCategoryDataSource xAxisValue = null;
    XDDFValueAxis yAxisPosition = null;       // y轴
    XDDFCategoryDataSource yAxisValue = null;
    if (chartDefaultOption.getJSONObject("axis") != null) {
      JSONObject xAxisUp = chartDefaultOption.getJSONObject("axis").getJSONObject("xAxisUp");
      JSONObject xAxisDown = chartDefaultOption.getJSONObject("axis").getJSONObject("xAxisDown");
      JSONObject yAxisLeft = chartDefaultOption.getJSONObject("axis").getJSONObject("yAxisLeft");
      JSONObject yAxisRight = chartDefaultOption.getJSONObject("axis").getJSONObject("yAxisRight");
      // x轴
      if (xAxisUp.getBoolean("show")) {
        xAxisPosition = chart.createCategoryAxis(AxisPosition.TOP);
        if (xAxisUp.getJSONObject("title").getBoolean("showTitle")) {
          xAxisPosition.setTitle(xAxisUp.getJSONObject("title").getString("text"));
        }
        if (xAxisUp.containsKey("data")) {
          xAxisValue = XDDFDataSourcesFactory.fromArray(
              Utils.JsonArray2ArrayString(xAxisUp.getJSONArray("data")));
        }
      }
      if (xAxisDown.getBoolean("show")) {
        xAxisPosition = chart.createCategoryAxis(AxisPosition.BOTTOM);
        if (xAxisDown.getJSONObject("title").getBoolean("showTitle")) {
          xAxisPosition.setTitle(xAxisDown.getJSONObject("title").getString("text"));
        }
        if (xAxisDown.containsKey("data")) {
          xAxisValue = XDDFDataSourcesFactory.fromArray(
              Utils.JsonArray2ArrayString(xAxisDown.getJSONArray("data")));
        }
      }
      // y轴
      if (yAxisLeft.getBoolean("show")) {
        yAxisPosition = chart.createValueAxis(AxisPosition.LEFT);
        if (yAxisLeft.getJSONObject("title").getBoolean("showTitle")) {
          yAxisPosition.setTitle(yAxisLeft.getJSONObject("title").getString("text"));
        }
        if (yAxisLeft.containsKey("data")) {
          yAxisValue = XDDFDataSourcesFactory.fromArray(
              Utils.JsonArray2ArrayString(yAxisLeft.getJSONArray("data")));
        }
      }
      if (yAxisRight.getBoolean("show")) {
        yAxisPosition = chart.createValueAxis(AxisPosition.RIGHT);
        if (yAxisRight.getJSONObject("title").getBoolean("showTitle")) {
          yAxisPosition.setTitle(yAxisRight.getJSONObject("title").getString("text"));
        }
        if (yAxisRight.containsKey("data")) {
          yAxisValue = XDDFDataSourcesFactory.fromArray(
              Utils.JsonArray2ArrayString(yAxisRight.getJSONArray("data")));
        }
      }
    }

    // 绘图：折线图，
    XDDFLineChartData draw_data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, xAxisPosition, yAxisPosition);

    if (chartDefaultOption.getJSONArray("series") != null) {
      JSONArray seriesData = chartDefaultOption.getJSONArray("series");

      // 逐个添加折线
      for (int series_idx = 0; series_idx < seriesData.size(); series_idx++) {
        JSONObject series_data = seriesData.getJSONObject(series_idx);

        // 1条折线的数据
        XDDFNumericalDataSource<Double> series_plot_data = XDDFDataSourcesFactory.fromArray(
            Utils.JsonArray2ArrayDouble(series_data.getJSONArray("data")));
        // 生成 1条折线
        XDDFLineChartData.Series series_plot = (XDDFLineChartData.Series) draw_data.addSeries(xAxisValue, series_plot_data);
        // 折线的标题
        series_plot.setTitle(series_data.getString("name"), null);
        if (series_data.getString("type") != null) {
          if (series_data.getString("type").equalsIgnoreCase("line")) {
            series_plot.setSmooth(false);              // 折线样式---直线

          }
        }
      }
    }
    chart.plot(draw_data);

    return sheet;
  }


  private static XSSFSheet getPieChart(XSSFSheet sheet, JSONObject chartData) {

    JSONObject chartOptions = chartData.getJSONObject("chartOptions");
    JSONObject chartDefaultOption = chartOptions.getJSONObject("defaultOption");

    // 创建一个画布
    XSSFDrawing drawing = sheet.createDrawingPatriarch();

    // 像素位置转换成 行列位置
    List<Integer> anchor_int = Utils.Pix2Anchor(
        chartData.getInteger("width"),
        chartData.getInteger("height"),
        chartData.getInteger("left"),
        chartData.getInteger("top"));
    // 创建一个chart对象
    XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0,
        anchor_int.get(0), anchor_int.get(1), anchor_int.get(2), anchor_int.get(3)));


    // **********************************  标题 轴 图例 配置 ***********************
    // 对defaultOption进行解析与设置

    // 标题
    if (chartDefaultOption.getJSONObject("title").getBoolean("show")) {
      // 获取title配置
      JSONObject chartTitle = chartDefaultOption.getJSONObject("title");
      // title内容
      chart.setTitleText(chartTitle.getString("text"));
      //        chart.setTitleOverlay(true);        // 标题覆盖
      // TODO 待完善
    }

    // 子标题  TODO 暂未找到如何设置
    if (chartDefaultOption.getJSONObject("subtitle").getBoolean("show")) {
      // 获取子标题配置
      JSONObject chartSubtitle = chartDefaultOption.getJSONObject("subtitle");
      // TODO 待完善
    }

    // 图例
    if (chartDefaultOption.getJSONObject("legend").getBoolean("show")) {
      // 获取图例配置
      JSONObject chartLegend = chartDefaultOption.getJSONObject("legend");
      // 图例位置
      XDDFChartLegend legend = chart.getOrAddLegend();
      String position = chartLegend.getJSONObject("position").getString("value");
      if (position.equalsIgnoreCase("right")) {
        legend.setPosition(LegendPosition.RIGHT);
      } else if (position.equalsIgnoreCase("left")) {
        legend.setPosition(LegendPosition.LEFT);
      } else if (position.equalsIgnoreCase("BOTTOM")) {
        legend.setPosition(LegendPosition.BOTTOM);
      } else {
        legend.setPosition(LegendPosition.TOP);
      }
      // TODO 待完善
    }

    //分类轴标数据，
    XDDFDataSource<String> xAxisValue = null;
    if (chartDefaultOption.getJSONObject("axis") != null) {
      JSONObject xAxisUp = chartDefaultOption.getJSONObject("axis").getJSONObject("xAxisUp");
      JSONObject xAxisDown = chartDefaultOption.getJSONObject("axis").getJSONObject("xAxisDown");
      JSONObject yAxisLeft = chartDefaultOption.getJSONObject("axis").getJSONObject("yAxisLeft");
      JSONObject yAxisRight = chartDefaultOption.getJSONObject("axis").getJSONObject("yAxisRight");
      // x轴
      if (xAxisUp.getBoolean("show")) {
        if (xAxisUp.containsKey("data")) {
          xAxisValue = XDDFDataSourcesFactory.fromArray(
              Utils.JsonArray2ArrayString(xAxisUp.getJSONArray("data")));
        }
      }
      if (xAxisDown.getBoolean("show")) {
        if (xAxisDown.containsKey("data")) {
          xAxisValue = XDDFDataSourcesFactory.fromArray(
              Utils.JsonArray2ArrayString(xAxisDown.getJSONArray("data")));
        }
      }
    }

    XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromArray(new Double[]{1.0,2.0,3.0,4.0,5.0,6.0,7.0,8.0,9.0,10.0,11.0});


    // **********************************  绘图 ***********************
    // 绘图，
    XDDFChartData draw_data = chart.createData(ChartTypes.PIE, null, null);
    //设置为可变颜色
    draw_data.setVaryColors(true);

    if (chartDefaultOption.getJSONArray("series") != null) {
      JSONArray seriesData = chartDefaultOption.getJSONArray("seriesData");
      for (int series_idx = 0; series_idx < seriesData.size(); series_idx++) {
        XDDFNumericalDataSource<Double> series_plot_data = XDDFDataSourcesFactory.fromArray(
            Utils.JsonArray2ArrayDouble(seriesData.getJSONArray(series_idx)));
        draw_data.addSeries(xAxisValue, series_plot_data);
      }
    }


    chart.plot(draw_data);

    return sheet;
  }


  private static XSSFSheet getColumnChart(XSSFSheet sheet, JSONObject chartData) {

    JSONObject chartOptions = chartData.getJSONObject("chartOptions");
    JSONObject chartDefaultOption = chartOptions.getJSONObject("defaultOption");

    // 创建一个画布
    XSSFDrawing drawing = sheet.createDrawingPatriarch();

    // 像素位置转换成 行列位置
    List<Integer> anchor_int = Utils.Pix2Anchor(
        chartData.getInteger("width"),
        chartData.getInteger("height"),
        chartData.getInteger("left"),
        chartData.getInteger("top"));
    // 创建一个chart对象
    XSSFChart chart = drawing.createChart(drawing.createAnchor(0, 0, 0, 0,
        anchor_int.get(0), anchor_int.get(1), anchor_int.get(2), anchor_int.get(3)));

    // **********************************  标题 轴 图例 配置 ***********************
    // 对defaultOption进行解析与设置

    // 标题
    if (chartDefaultOption.getJSONObject("title").getBoolean("show")) {
      // 获取title配置
      JSONObject chartTitle = chartDefaultOption.getJSONObject("title");
      // title内容
      chart.setTitleText(chartTitle.getString("text"));
      //        chart.setTitleOverlay(true);        // 标题覆盖
      // TODO 待完善
    }

    // 子标题  TODO 暂未找到如何设置
    if (chartDefaultOption.getJSONObject("subtitle").getBoolean("show")) {
      // 获取子标题配置
      JSONObject chartSubtitle = chartDefaultOption.getJSONObject("subtitle");
      // TODO 待完善
    }

    // 图例
    if (chartDefaultOption.getJSONObject("legend").getBoolean("show")) {
      // 获取图例配置
      JSONObject chartLegend = chartDefaultOption.getJSONObject("legend");
      // 图例位置
      XDDFChartLegend legend = chart.getOrAddLegend();
      String position = chartLegend.getJSONObject("position").getString("value");
      if (position.equalsIgnoreCase("right")) {
        legend.setPosition(LegendPosition.RIGHT);
      } else if (position.equalsIgnoreCase("left")) {
        legend.setPosition(LegendPosition.LEFT);
      } else if (position.equalsIgnoreCase("BOTTOM")) {
        legend.setPosition(LegendPosition.BOTTOM);
      } else {
        legend.setPosition(LegendPosition.TOP);
      }
      // TODO 待完善
    }

    // 坐标轴   TODO 样式设置
    XDDFCategoryAxis xAxisPosition = null;    // x轴
    XDDFCategoryDataSource xAxisValue = null;
    XDDFValueAxis yAxisPosition = null;       // y轴
    XDDFCategoryDataSource yAxisValue = null;
    if (chartDefaultOption.getJSONObject("axis") != null) {
      JSONObject xAxisUp = chartDefaultOption.getJSONObject("axis").getJSONObject("xAxisUp");
      JSONObject xAxisDown = chartDefaultOption.getJSONObject("axis").getJSONObject("xAxisDown");
      JSONObject yAxisLeft = chartDefaultOption.getJSONObject("axis").getJSONObject("yAxisLeft");
      JSONObject yAxisRight = chartDefaultOption.getJSONObject("axis").getJSONObject("yAxisRight");
      // x轴
      if (xAxisUp.getBoolean("show")) {
        xAxisPosition = chart.createCategoryAxis(AxisPosition.TOP);
        if (xAxisUp.getJSONObject("title").getBoolean("showTitle")) {
          xAxisPosition.setTitle(xAxisUp.getJSONObject("title").getString("text"));
        }
        if (xAxisUp.containsKey("data")) {
          xAxisValue = XDDFDataSourcesFactory.fromArray(
              Utils.JsonArray2ArrayString(xAxisUp.getJSONArray("data")));
        }
      }
      if (xAxisDown.getBoolean("show")) {
        xAxisPosition = chart.createCategoryAxis(AxisPosition.BOTTOM);
        if (xAxisDown.getJSONObject("title").getBoolean("showTitle")) {
          xAxisPosition.setTitle(xAxisDown.getJSONObject("title").getString("text"));
        }
        if (xAxisDown.containsKey("data")) {
          xAxisValue = XDDFDataSourcesFactory.fromArray(
              Utils.JsonArray2ArrayString(xAxisDown.getJSONArray("data")));
        }
      }
      // y轴
      if (yAxisLeft.getBoolean("show")) {
        yAxisPosition = chart.createValueAxis(AxisPosition.LEFT);
        if (yAxisLeft.getJSONObject("title").getBoolean("showTitle")) {
          yAxisPosition.setTitle(yAxisLeft.getJSONObject("title").getString("text"));
        }
        if (yAxisLeft.containsKey("data")) {
          yAxisValue = XDDFDataSourcesFactory.fromArray(
              Utils.JsonArray2ArrayString(yAxisLeft.getJSONArray("data")));
        }
      }
      if (yAxisRight.getBoolean("show")) {
        yAxisPosition = chart.createValueAxis(AxisPosition.RIGHT);
        if (yAxisRight.getJSONObject("title").getBoolean("showTitle")) {
          yAxisPosition.setTitle(yAxisRight.getJSONObject("title").getString("text"));
        }
        if (yAxisRight.containsKey("data")) {
          yAxisValue = XDDFDataSourcesFactory.fromArray(
              Utils.JsonArray2ArrayString(yAxisRight.getJSONArray("data")));
        }
      }
    }

    // 绘图：折线图，
    XDDFLineChartData draw_data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, xAxisPosition, yAxisPosition);

    if (chartDefaultOption.getJSONArray("series") != null) {
      JSONArray seriesData = chartDefaultOption.getJSONArray("series");

      // 逐个添加折线
      for (int series_idx = 0; series_idx < seriesData.size(); series_idx++) {
        JSONObject series_data = seriesData.getJSONObject(series_idx);

        // 1条折线的数据
        XDDFNumericalDataSource<Double> series_plot_data = XDDFDataSourcesFactory.fromArray(
            Utils.JsonArray2ArrayDouble(series_data.getJSONArray("data")));
        // 生成 1条折线
        XDDFLineChartData.Series series_plot = (XDDFLineChartData.Series) draw_data.addSeries(xAxisValue, series_plot_data);
        // 折线的标题
        series_plot.setTitle(series_data.getString("name"), null);
        if (series_data.getString("type") != null) {
          if (series_data.getString("type").equalsIgnoreCase("line")) {
            series_plot.setSmooth(false);              // 折线样式---直线

          }
        }
      }
    }
    chart.plot(draw_data);

    return sheet;
  }

}
