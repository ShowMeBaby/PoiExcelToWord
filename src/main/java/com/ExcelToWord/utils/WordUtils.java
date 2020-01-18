package com.ExcelToWord.utils;


import com.ExcelToWord.entity.CellEntity;
import com.ExcelToWord.entity.VarMapEntity;
import com.ExcelToWord.utils.ExcelUtils;
import com.alibaba.fastjson.JSON;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;


public class WordUtils {
    /**
     * 导入成功后返回的对象
     */
    public static class ReplaceDocxResult {
        private boolean success;
        private String message;
        private ArrayList<String> undefinedVar;
        private Integer keySize;

        public ReplaceDocxResult(boolean success, String message, ArrayList<String> undefinedVar, Integer keySize) {
            this.success = success;
            this.message = message;
            this.undefinedVar = undefinedVar;
            this.keySize = keySize;
        }

        public boolean isSuccess() {
            return success;
        }

        public String getMessage() {
            return message;
        }

        public ArrayList<String> getUndefinedVar() {
            return undefinedVar;
        }

        public Integer getKeySize() {
            return keySize;
        }

        public String toJson() {
            return JSON.toJSONString(this);
        }
    }

    /**
     * 替换Docx文档中的变量名为传入变量
     *
     * @param xlsxPath 变量表xlsx的文档路径
     * @param inputSrc 需要替换的word文档路径
     * @param outSrc   替换成功后的word文档存储路径
     * @return 操作是否成功
     */
    public static ReplaceDocxResult replaceDocxText(String xlsxPath, String inputSrc, String outSrc) {
        try {
            Workbook workbook = ExcelUtils.getWorkbook(xlsxPath);
            Map<String, VarMapEntity> map = ExcelUtils.getSheetVarMap(workbook);
            boolean success = false;
            ArrayList<String> undefinedVar = new ArrayList<>();
            try {
                XWPFDocument docx = new XWPFDocument(POIXMLDocument.openPackage(inputSrc));
                /* 替换段落中指定的文本 */
                for (XWPFParagraph p : docx.getParagraphs()) {
                    List<XWPFRun> runs = p.getRuns();
                    if (runs != null) {
                        for (XWPFRun r : runs) {
                            //需要替换的文本
                            String text = r.getText(0);
                            ArrayList<String> strVarsList = getStrVarsList(text);
                            for (String varName : strVarsList) {
                                if (map.containsKey(varName)) {
                                    String varType = map.get(varName).getType();
                                    if ("String".equals(varType)) {
                                        text = text.replace(varName, map.get(varName).getStringValue());
                                        //0是替换全部，如果不设置那么默认就是从原文字结尾开始追加
                                        r.setText(text, 0);
                                    } else if ("Table".equals(varType)) {
                                        // 清空标记
                                        r.setText("", 0);
                                        // 创建表格
                                        createWordTable(p, docx, map.get(varName).getTableValue());
                                    }
                                } else {
                                    // 匹配到word文档需要,但变量表中不存在的变量
                                    undefinedVar.add(varName);
                                }
                            }
                        }
                    }
                }
                /* 替换表格中指定的文字 */
                for (XWPFTable tab : docx.getTables()) {
                    for (XWPFTableRow row : tab.getRows()) {
                        for (XWPFTableCell cell : row.getTableCells()) {
                            //注意，getParagraphs一定不能漏掉
                            //因为一个表格里面可能会有多个需要替换的文字
                            //如果没有这个步骤那么文字会替换不了
                            for (XWPFParagraph p : cell.getParagraphs()) {
                                for (XWPFRun r : p.getRuns()) {
                                    String text = r.getText(0);
                                    ArrayList<String> strVarsList = getStrVarsList(text);
                                    for (String varName : strVarsList) {
                                        if (map.containsKey(varName)) {
                                            // 表格中不能插入表格,直接取字符串值取不到就拉倒
                                            text = text.replace(varName, map.get(varName).getStringValue());
                                            r.setText(text, 0);
                                        } else {
                                            // 匹配到word文档需要,但变量表中不存在的变量
                                            undefinedVar.add(varName);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                docx.write(new FileOutputStream(outSrc));
                success = true;
            } catch (IOException ignored) {
            }
            return new ReplaceDocxResult(success, undefinedVar.size() > 0 ? "解析成功!但文档中使用到了变量表未定义的变量,请核实." : "解析成功!", undefinedVar,map.size());
        } catch (IOException e) {
            return new ReplaceDocxResult(false, "解析失败!文档不存在!", new ArrayList<>(),0);
        } catch (Exception e) {
            return new ReplaceDocxResult(false, "解析失败!变量表不存在!", new ArrayList<>(),0);
        }
    }

    /**
     * 向文档中插入表格
     *
     * @param xWPFParagraph 段落
     * @param docx          文档对象
     * @param tableValue    需要插入的表格数据
     */
    public static void createWordTable(XWPFParagraph xWPFParagraph, XWPFDocument docx, ArrayList<ArrayList<CellEntity>> tableValue) {
        // 插入表格
        XmlCursor cursor = xWPFParagraph.getCTP().newCursor();
        XWPFTable table = docx.insertNewTbl(cursor);
        // 设置表格宽
        table.setWidth("9000");
        // 设置边框加粗
        CTTblBorders borders = table.getCTTbl().getTblPr().getTblBorders();
        borders.getTop().setSz(BigInteger.valueOf(12));
        borders.getLeft().setSz(BigInteger.valueOf(12));
        borders.getRight().setSz(BigInteger.valueOf(12));
        borders.getBottom().setSz(BigInteger.valueOf(12));
        // 遍历表格数据
        for (int i = 0; i < tableValue.size(); i++) {
            ArrayList<CellEntity> cellEntities = tableValue.get(i);
            List<String> collect = cellEntities.stream().filter(o -> !o.getValue().isEmpty()).map(CellEntity::getValue).collect(Collectors.toList());
            // 如果整行都没数据,则过滤掉
            if (collect.size() > 0) {
                if (0 == i) {
                    XWPFTableRow row_0 = table.getRow(0);
                    for (int j = 0; j < cellEntities.size(); j++) {
                        CellEntity cellEntity = cellEntities.get(j);
                        if (0 == j) {
                            row_0.getCell(0).setText(cellEntity.getValue());
                        } else {
                            row_0.addNewTableCell().setText(cellEntity.getValue());
                        }
                    }
                } else {
                    XWPFTableRow row_1 = table.createRow();
                    for (int j = 0; j < cellEntities.size(); j++) {
                        CellEntity cellEntity = cellEntities.get(j);
                        row_1.getCell(j).setText(cellEntity.getValue());
                    }
                }
            }
        }
        setTableLocation(table, "center");
        setCellLocation(table, "CENTER", "center");
    }

    /**
     * 找出字符串中所有的变量列表
     *
     * @param regString 所要查询的字符串
     * @return 查询到的变量数组
     */
    public static ArrayList<String> getStrVarsList(String regString) {
        ArrayList<String> result = new ArrayList<>();
        if (regString != null) {
            Matcher matcher = Pattern.compile("\\$\\{[\\u4e00-\\u9fa5\\w]+}").matcher(regString);
            while (matcher.find()) {
                result.add(matcher.group());
            }
        }
        return result;
    }

    /**
     * 设置单元格水平位置和垂直位置
     *
     * @param xwpfTable
     * @param verticalLoction    单元格中内容垂直上TOP，下BOTTOM，居中CENTER，BOTH两端对齐
     * @param horizontalLocation 单元格中内容水平居中center,left居左，right居右，both两端对齐
     */
    private static void setCellLocation(XWPFTable xwpfTable, String verticalLoction, String horizontalLocation) {
        List<XWPFTableRow> rows = xwpfTable.getRows();
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                CTTc cttc = cell.getCTTc();
                CTP ctp = cttc.getPList().get(0);
                CTPPr ctppr = ctp.getPPr();
                if (ctppr == null) {
                    ctppr = ctp.addNewPPr();
                }
                CTJc ctjc = ctppr.getJc();
                if (ctjc == null) {
                    ctjc = ctppr.addNewJc();
                }
                ctjc.setVal(STJc.Enum.forString(horizontalLocation)); //水平居中
                cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.valueOf(verticalLoction));//垂直居中
            }
        }
    }

    /**
     * 设置表格位置
     *
     * @param xwpfTable
     * @param location  整个表格居中center,left居左，right居右，both两端对齐
     */
    private static void setTableLocation(XWPFTable xwpfTable, String location) {
        CTTbl cttbl = xwpfTable.getCTTbl();
        CTTblPr tblpr = cttbl.getTblPr() == null ? cttbl.addNewTblPr() : cttbl.getTblPr();
        CTJc cTJc = tblpr.addNewJc();
        cTJc.setVal(STJc.Enum.forString(location));
    }
}