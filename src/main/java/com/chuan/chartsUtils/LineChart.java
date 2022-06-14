package com.chuan.chartsUtils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.PresetLineDash;
import org.apache.poi.xddf.usermodel.XDDFLineProperties;
import org.apache.poi.xddf.usermodel.XDDFPresetLineDash;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.util.Units;

import java.awt.*;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class LineChart {
  private static final int Row_Height_Pix = 18;
  private static final int Col_Width_Pix = 61;

  private static XSSFWorkbook getSheet(){
    XSSFWorkbook wb = new XSSFWorkbook();
    String sheetName = "chart";
    XSSFSheet sheet = wb.createSheet(sheetName);
    // 第一行，国家名称
    Row row = sheet.createRow(0);
    Cell cell = row.createCell(0);
    cell.setCellValue("俄罗斯");
    cell = row.createCell(1);
    cell.setCellValue("加拿大");
    cell = row.createCell(2);
    cell.setCellValue("美国");
    cell = row.createCell(3);
    cell.setCellValue("中国");
    cell = row.createCell(4);
    cell.setCellValue("巴西");
    cell = row.createCell(5);
    cell.setCellValue("澳大利亚");
    cell = row.createCell(6);
    cell.setCellValue("印度");
    // 第二行，乡村地区
    row = sheet.createRow(1);
    cell = row.createCell(0);
    cell.setCellValue(17098242);
    cell = row.createCell(1);
    cell.setCellValue(9984670);
    cell = row.createCell(2);
    cell.setCellValue(9826675);
    cell = row.createCell(3);
    cell.setCellValue(9596961);
    cell = row.createCell(4);
    cell.setCellValue(8514877);
    cell = row.createCell(5);
    cell.setCellValue(7741220);
    cell = row.createCell(6);
    cell.setCellValue(3287263);
    // 第三行，农村人口
    row = sheet.createRow(2);
    cell = row.createCell(0);
    cell.setCellValue(14590041);
    cell = row.createCell(1);
    cell.setCellValue(35151728);
    cell = row.createCell(2);
    cell.setCellValue(32993302);
    cell = row.createCell(3);
    cell.setCellValue(14362887);
    cell = row.createCell(4);
    cell.setCellValue(21172141);
    cell = row.createCell(5);
    cell.setCellValue(25335727);
    cell = row.createCell(6);
    cell.setCellValue(13724923);
    // 第四行，面积平局
    row = sheet.createRow(3);
    cell = row.createCell(0);
    cell.setCellValue(9435701.143);
    cell = row.createCell(1);
    cell.setCellValue(9435701.143);
    cell = row.createCell(2);
    cell.setCellValue(9435701.143);
    cell = row.createCell(3);
    cell.setCellValue(9435701.143);
    cell = row.createCell(4);
    cell.setCellValue(9435701.143);
    cell = row.createCell(5);
    cell.setCellValue(9435701.143);
    cell = row.createCell(6);
    cell.setCellValue(9435701.143);
    // 第四行，人口平局
    row = sheet.createRow(4);
    cell = row.createCell(0);
    cell.setCellValue(22475821.29);
    cell = row.createCell(1);
    cell.setCellValue(22475821.29);
    cell = row.createCell(2);
    cell.setCellValue(22475821.29);
    cell = row.createCell(3);
    cell.setCellValue(22475821.29);
    cell = row.createCell(4);
    cell.setCellValue(22475821.29);
    cell = row.createCell(5);
    cell.setCellValue(22475821.29);
    cell = row.createCell(6);
    cell.setCellValue(22475821.29);

    return wb;

  }

  public static void getLineChar() throws IOException {
    FileOutputStream fileOut = null;
    XSSFWorkbook wb = null;
    try {
      wb = getSheet();
      XSSFSheet sheet = wb.getSheetAt(0);
      // 创建一个画布
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      // 前四个默认0，[0,5]：从0列5行开始;[7,26]:到7列26行结束
      // 默认宽度(14-8)*12
      int col1 = 150 / Col_Width_Pix;    // left
      int row1 = 450 / Row_Height_Pix;   // top
      int col2 = (600 + 150) / Col_Width_Pix;  // width+left
      int row2 = (250 + 450) / Row_Height_Pix; // hight+top
      XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, col1, row1, col2, row2);
      // 创建一个chart对象
      XSSFChart chart = drawing.createChart(anchor);
      // 标题
      chart.setTitleText("地区排名前七的国家");
      // 标题覆盖
      chart.setTitleOverlay(false);

      // 图例位置
      XDDFChartLegend legend = chart.getOrAddLegend();
      legend.setPosition(LegendPosition.TOP);

      // 分类轴标(X轴),标题位置
      XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
      bottomAxis.setTitle("国家");
      // 值(Y轴)轴,标题位置
      XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
      leftAxis.setTitle("面积和人口");

      // CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）
      // 分类轴标(X轴)数据，单元格范围位置[0, 0]到[0, 6]
      XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, 6));
      // XDDFCategoryDataSource countries = XDDFDataSourcesFactory.fromArray(new String[] {"俄罗斯","加拿大","美国","中国","巴西","澳大利亚","印度"});
      // 数据1，单元格范围位置[1, 0]到[1, 6]
      XDDFNumericalDataSource<Double> area = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 6));
      // XDDFNumericalDataSource<Integer> area = XDDFDataSourcesFactory.fromArray(new Integer[] {17098242,9984670,9826675,9596961,8514877,7741220,3287263});

      // 数据1，单元格范围位置[2, 0]到[2, 6]
      XDDFNumericalDataSource<Double> population = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(2, 2, 0, 6));

      // LINE：折线图，
      XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);

      // 图表加载数据，折线1
      XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(countries, area);
      // 折线图例标题
      series1.setTitle("面积", null);
      // 直线
      series1.setSmooth(false);
      // 设置标记大小
      series1.setMarkerSize((short) 6);
      // 设置标记样式，星星
      series1.setMarkerStyle(MarkerStyle.STAR);

      // 图表加载数据，折线2
      XDDFLineChartData.Series series2 = (XDDFLineChartData.Series) data.addSeries(countries, population);
      // 折线图例标题
      series2.setTitle("人口", null);
      // 曲线
      series2.setSmooth(true);
      // 设置标记大小
      series2.setMarkerSize((short) 6);
      // 设置标记样式，正方形
      series2.setMarkerStyle(MarkerStyle.SQUARE);

      // 图表加载数据，平均线3
      // 数据1，单元格范围位置[2, 0]到[2, 6]
      XDDFNumericalDataSource<Double> population3 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(3, 3, 0, 6));
      XDDFLineChartData.Series series3 = (XDDFLineChartData.Series) data.addSeries(countries, population3);
      // 折线图例标题
      series3.setTitle("面积平均", null);
      // 直线
      series3.setSmooth(false);
      // 设置标记大小
      // series3.setMarkerSize((short) 3);
      // 设置标记样式，正方形
      series3.setMarkerStyle(MarkerStyle.NONE);
      // 折线图LineChart
      // XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.CHARTREUSE));
      XDDFLineProperties line = new XDDFLineProperties();
      // line.setFillProperties(fill);
      // line.setLineCap(LineCap.ROUND);
      line.setPresetDash(new XDDFPresetLineDash(PresetLineDash.DOT));// 虚线
      // XDDFShapeProperties shapeProperties = new XDDFShapeProperties();
      // shapeProperties.setLineProperties(line);
      // series3.setShapeProperties(shapeProperties);
      series3.setLineProperties(line);

      // 图表加载数据，平均线3
      // 数据1，单元格范围位置[2, 0]到[2, 6]
      XDDFNumericalDataSource<Double> population4 = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(4, 4, 0, 6));
      XDDFLineChartData.Series series4 = (XDDFLineChartData.Series) data.addSeries(countries, population4);
      // 折线图例标题
      series4.setTitle("人口平均", null);
      // 直线
      series4.setSmooth(false);
      // 设置标记大小
      // series4.setMarkerSize((short) 3);
      // 设置标记样式，正方形
      series4.setMarkerStyle(MarkerStyle.NONE);
      XDDFLineProperties line4 = new XDDFLineProperties();
      line4.setPresetDash(new XDDFPresetLineDash(PresetLineDash.DOT));// 虚线
      series4.setLineProperties(line);

      // 绘制
      chart.plot(data);

      // 打印图表的xml
      // System.out.println(chart.getCTChart());

      // 将输出写入excel文件
      String filename = "排行榜前七的国家.xlsx";
      try {
        fileOut = new FileOutputStream(filename);
      } catch (FileNotFoundException ex) {
        ex.printStackTrace();
      }
      try {
        wb.write(fileOut);
      } catch (IOException ex) {
        ex.printStackTrace();
      }
    } catch (Exception e) {
      e.printStackTrace();
    } finally {
      wb.close();
      if (fileOut != null) {
        try {
          fileOut.close();
        } catch (IOException ex) {
          ex.printStackTrace();
        }
      }
    }
  }

  public static void main(String[] args) throws IOException {
    getLineChar();
  }
}
