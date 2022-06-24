package com.chuan.chartsUtils;

import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;


public class DataVerificationFun {

  public static void setCellVerfication(XSSFWorkbook workbook, XSSFSheet sheet, JSONObject verficationObject, String fontFormat) {
    for (String key : verficationObject.keySet()) {
      int row_idx = Integer.parseInt(key.split("_")[0]);
      int col_idx = Integer.parseInt(key.split("_")[1]);

      JSONObject thisVerfiObject = verficationObject.getJSONObject(key);
      String verifiType = thisVerfiObject.getString("type");
      switch (verifiType) {
        case "dropdown":   // 下拉列表
          setDropDownList(workbook, sheet, row_idx, col_idx, thisVerfiObject);
          break;
        case "checkbox":  // 复选框
          // 暂无法处理
          break;
        case "number": // 数字
        case "number_integer": //(数字-整数) ;
        case "number_decimal":  // (数字-小数)
        case "text_length":   // (文本-长度)
        case "date": // (日期)
          setNumberCheck(sheet, row_idx, col_idx, thisVerfiObject, fontFormat);
          break;
        case "text_content":    //(文本-内容)
        case "validity":    //(有效性)
          // 暂无法处理
          break;
      }
    }
  }

  private static void setDropDownList(XSSFWorkbook wb, XSSFSheet sheet, int row_idx, int col_idx, JSONObject verifiObject) {
    // 下拉候选数据
    String optionsStr = verifiObject.getString("value1");
    String[] options = optionsStr.split(",");

    if (optionsStr.length() > 255) {
      //获取所有sheet页个数
      int sheetTotal = wb.getNumberOfSheets();
      String hiddenSheetName = "hiddenSheet" + sheetTotal;
      XSSFSheet hiddenSheet = wb.createSheet(hiddenSheetName);
      Row row;
      //写入下拉数据到新的sheet页中
      for (int i = 0; i < options.length; i++) {
        row = hiddenSheet.createRow(i);
        Cell cell = row.createCell(0);
        cell.setCellValue(options[i]);
      }
      //获取新sheet页内容
      String strFormula = hiddenSheetName + "!$A$1:$A" + "$" + options.length + 1;
      XSSFDataValidationConstraint constraint = new XSSFDataValidationConstraint(DataValidationConstraint.ValidationType.LIST, strFormula);
      // 设置数据有效性加载在哪个单元格上
      CellRangeAddressList regions = new CellRangeAddressList(row_idx, row_idx, col_idx, col_idx);
      // 数据有效性对象
      DataValidationHelper help = new XSSFDataValidationHelper(sheet);
      DataValidation validation = help.createValidation(constraint, regions);
      sheet.addValidationData(validation);
      //将新建的sheet页隐藏掉
      wb.setSheetHidden(sheetTotal, true);
    } else {
      XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
      XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
          .createExplicitListConstraint(options);
      CellRangeAddressList addressList = new CellRangeAddressList(row_idx, row_idx, col_idx, col_idx);
      XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, addressList);

      // 选中单元格时显示提示语
      if (verifiObject.getBoolean("hintShow")) {
        // 提示语文本
        validation.createPromptBox(null, verifiObject.getString("hintText"));
        validation.setShowPromptBox(true);
      }
      sheet.addValidationData(validation);
    }
  }

  private static void setNumberCheck(XSSFSheet sheet, int row_idx, int col_idx, JSONObject verifiObject, String fontFormat) {

    String conditionValue1 = verifiObject.getString("value1");
    String conditionValue2 = verifiObject.getString("value2");
    String verifiType = verifiObject.getString("type");
    int validationType;
    switch (verifiType) {
      case "number_integer":
        validationType = DataValidationConstraint.ValidationType.INTEGER;
        break;
      case "text_length":
        validationType = DataValidationConstraint.ValidationType.TEXT_LENGTH;
        break;
      case "date":
        validationType = DataValidationConstraint.ValidationType.DATE;
        break;
      default:
        validationType = DataValidationConstraint.ValidationType.DECIMAL;
    }
    // 条件类型 "bw"(介于)，"nb"(不介于)，"eq"(等于)，"ne"(不等于)，"gt"(大于)，"lt"(小于), "gte"(大于等于)，"lte"(小于等于)
    //    "bf"(早于) "nbf"(不早于) "af"(晚于) "naf"(不晚于)
    String conditionType = verifiObject.getString("type2");
    int operatorType;
    switch (conditionType) {
      case "bw":
        operatorType = DataValidationConstraint.OperatorType.BETWEEN;
        break;
      case "nb":
        operatorType = DataValidationConstraint.OperatorType.NOT_BETWEEN;
        break;
      case "eq":
        operatorType = DataValidationConstraint.OperatorType.EQUAL;
        break;
      case "ne":
        operatorType = DataValidationConstraint.OperatorType.NOT_EQUAL;
        break;
      case "gt":
      case "af":
        operatorType = DataValidationConstraint.OperatorType.GREATER_THAN;
        break;
      case "gte":
      case "naf":
        operatorType = DataValidationConstraint.OperatorType.GREATER_OR_EQUAL;
        break;
      case "lte":
      case "nbf":
        operatorType = DataValidationConstraint.OperatorType.LESS_OR_EQUAL;
        break;
      case "lt":
      case "bf":
        operatorType = DataValidationConstraint.OperatorType.LESS_THAN;
        break;
      default:
        operatorType = DataValidationConstraint.OperatorType.IGNORED;
    }

    XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
    DataValidationConstraint dvConstraint;
    switch (validationType){
      case DataValidationConstraint.ValidationType.DECIMAL:
        dvConstraint = dvHelper.createDecimalConstraint(operatorType, conditionValue1, conditionValue2);
        break;
      case DataValidationConstraint.ValidationType.DATE:
        dvConstraint = dvHelper.createDateConstraint(operatorType, conditionValue1, conditionValue2, fontFormat);
        break;
      case DataValidationConstraint.ValidationType.INTEGER:
        dvConstraint = dvHelper.createIntegerConstraint(operatorType, conditionValue1, conditionValue2);
        break;
      case DataValidationConstraint.ValidationType.TEXT_LENGTH:
        dvConstraint = dvHelper.createTextLengthConstraint(operatorType,conditionValue1,conditionValue2);
        break;
      case DataValidationConstraint.ValidationType.TIME:
        dvConstraint = dvHelper.createTimeConstraint(operatorType,conditionValue1,conditionValue2);
        break;
      default:
        dvConstraint = dvHelper.createNumericConstraint(validationType, operatorType, conditionValue1, conditionValue2);

    }

    CellRangeAddressList addressList = new CellRangeAddressList(row_idx, row_idx, col_idx, col_idx);
    XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(dvConstraint, addressList);

    //错误提示信息
    validation.createErrorBox("错误提示", verifiType+"-"+conditionType+"-"+conditionValue1+"-"+conditionValue2);
    //设置是否显示错误窗口
    validation.setShowErrorBox(true);

    // 选中单元格时显示提示语
    if (verifiObject.getBoolean("hintShow")) {
      // 提示语文本
      validation.createPromptBox(null, verifiObject.getString("hintText"));
      validation.setShowPromptBox(true);
    }
    sheet.addValidationData(validation);
  }

}
