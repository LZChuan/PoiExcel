package com.chuan.chartsUtils;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Pie {

  public static void main(String[] args) throws IOException {
    XSSFWorkbook wb = new XSSFWorkbook();
    FileOutputStream fileOut = null;
    try {
      XSSFSheet sheet = wb.createSheet("Sheet1");
      //第一行，国家名称
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

      //创建一个画布
      XSSFDrawing drawing = sheet.createDrawingPatriarch();
      //前四个默认0，[0,4]：从0列4行开始;[7,20]:宽度7个单元格，20向下扩展到20行
      //默认宽度(14-8)*12
      XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 4, 7, 20);
      //创建一个chart对象
      XSSFChart chart = drawing.createChart(anchor);
      //标题
      chart.setTitleText("地区排名前七的国家");
      //标题是否覆盖图表
      chart.setTitleOverlay(false);

      //图例位置
      XDDFChartLegend legend = chart.getOrAddLegend();
      legend.setPosition(LegendPosition.TOP_RIGHT);

      //CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）
      //分类轴标数据，
      XDDFDataSource<String> countries = XDDFDataSourcesFactory.fromStringCellRange(sheet, new CellRangeAddress(0, 0, 0, 6));
      //数据1，
      XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(1, 1, 0, 6));
      //XDDFChartData data = chart.createData(ChartTypes.PIE3D, null, null);
      XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
      //设置为可变颜色
      data.setVaryColors(true);
      //图表加载数据
      data.addSeries(countries, values);

      //绘制
      chart.plot(data);

      // 将输出写入excel文件
      fileOut = new FileOutputStream("pie-chart-top-seven-countries.xlsx");
      wb.write(fileOut);
    } catch (Exception e) {
      e.printStackTrace();
    } finally {
      wb.close();
      fileOut.close();
    }
  }

}