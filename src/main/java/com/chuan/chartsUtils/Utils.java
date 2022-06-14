package com.chuan.chartsUtils;


import sun.java2d.loops.FontInfo;

public class Utils {

    private static final short TWIPS_PER_PIEXL = 15; //1 Pixel = 1440 TPI / 96 DPI = 15 Twips

    public static short pixel2PoiHeight(int pixel) {
      return (short) (pixel * TWIPS_PER_PIEXL);
    }

    public static int poiHeight2Pixel(short height) {
      return height / TWIPS_PER_PIEXL;
    }

//  https://blog.csdn.net/feg545/article/details/11983429
}
