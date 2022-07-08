package com.chuan.chartsUtils;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

public class PythonTest {

  public static void main(String[] args) {
    String executer = "python";
    // python绝对路径
    String file_path = "D:\\my_workspace\\code_project\\Java_Project\\PoiExcel\\src\\main\\java\\com\\chuan\\chartsUtils\\SparkLineProcess.py";
    String num1 = "3";
    String num2 = "7";
    String[] command_line = new String[] {executer, file_path, num1, num2};
    try {
      Process process = Runtime.getRuntime().exec(command_line);
      BufferedReader in = new BufferedReader(new InputStreamReader(process.getInputStream(), "GBK"));
      String line;
      while ((line = in.readLine()) != null) {
        System.out.println(line);
      }
      in.close();
      // java代码中的 process.waitFor() 返回值（和我们通常意义上见到的0与1定义正好相反）
      // 返回值为0 - 表示调用python脚本成功；
      // 返回值为1 - 表示调用python脚本失败。
      int re = process.waitFor();
      System.out.println("调用 python 脚本是否成功：" + re);
    } catch (IOException | InterruptedException e1) {
      e1.printStackTrace();
    }
  }
}
