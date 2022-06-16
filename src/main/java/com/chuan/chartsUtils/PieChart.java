package com.chuan.chartsUtils;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.util.List;

public class PieChart {
  private static final int Row_Height_Pix = 18;
  private static final int Col_Width_Pix = 61;

  /**
   * 创建饼图（xlsx格式excel）
   * @param sheetAt 工作表
   */
  private static void createPie(XSSFSheet sheetAt) {
    // 创建一个画布
    XSSFDrawing drawing = sheetAt.createDrawingPatriarch();
    //设置画布在excel工作表的位置
    List<Integer> anchor_int = Utils.Pix2Anchor(400, 250, 500, 120);
    XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0,
        anchor_int.get(0), anchor_int.get(1), anchor_int.get(2), anchor_int.get(3));
    // 创建一个chart对象
    XSSFChart chart = drawing.createChart(anchor);



    CTChart ctChart = chart.getCTChart();
    CTPlotArea ctPlotArea = ctChart.getPlotArea();
    // 创建圆环图
    CTPieChart ctPieChart = ctPlotArea.addNewPieChart();
    CTBoolean ctBoolean = ctPieChart.addNewVaryColors();
    // 允许自定义颜色
    ctBoolean.setVal(true);

    // 创建序列,并且设置选中区域
    CTPieSer ctPieSer = ctPieChart.addNewSer();

    // 数据区域
    CTNumDataSource ctNumDataSource = ctPieSer.addNewVal();
    CTNumRef ctNumRef = ctNumDataSource.addNewNumRef();
    // 选第1-6行,第1-3列作为数据区域 //1 2 3
    String numDataRange = new CellRangeAddress(1, 4, 1, 1).formatAsString(sheetAt.getSheetName(),
        true);
    ctNumRef.setF(numDataRange);


    // 设置标签格式
    CTDLbls newDLbls = ctPieSer.addNewDLbls();
    //		newDLbls.setShowLegendKey(ctBoolean);
    newDLbls.setShowVal(ctBoolean);
    //		newDLbls.setShowCatName(ctBoolean);//显示横坐标（图注）
    newDLbls.setShowPercent(ctBoolean);// 显示百分比
    newDLbls.setShowBubbleSize(ctBoolean);// 显示纵坐标（数量）
    newDLbls.setShowLeaderLines(ctBoolean);// 显示线


  }

}
