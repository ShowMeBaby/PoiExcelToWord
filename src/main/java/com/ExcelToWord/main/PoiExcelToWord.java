package com.ExcelToWord.main;

import com.ExcelToWord.utils.WordUtils;
public class PoiExcelToWord {
    public static void main(String[] args) {
        String xlsxPath = System.getProperty("xlsxPath");
        String inputPath = System.getProperty("inputPath");
        String outPath = System.getProperty("outPath");
        WordUtils.ReplaceDocxResult replaceDocxResult = WordUtils.replaceDocxText(xlsxPath, inputPath, outPath);
        System.out.println(replaceDocxResult.toJson());
    }
}
