package com.ExcelToWord.entity;
import java.util.Arrays;
public class CellEntity {
    private String value;
    private int startRowIndex;
    private int endRowIndex;
    private int startColumnIndex;
    private int endColumnIndex;
    private int[] mergedRegions;
    private boolean toMergedRegion; // 临时变量,只有合并单元格用到,用来判断是否重复填充合并单元格的值
    public CellEntity(String value, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex) {
        this.value = value;
        this.startRowIndex = startRowIndex;
        this.endRowIndex = endRowIndex;
        this.startColumnIndex = startColumnIndex;
        this.endColumnIndex = endColumnIndex;
        this.mergedRegions = new int[]{};
        this.toMergedRegion = false;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public int getStartRowIndex() {
        return startRowIndex;
    }

    public void setStartRowIndex(int startRowIndex) {
        this.startRowIndex = startRowIndex;
    }

    public int getEndRowIndex() {
        return endRowIndex;
    }

    public void setEndRowIndex(int endRowIndex) {
        this.endRowIndex = endRowIndex;
    }

    public int getStartColumnIndex() {
        return startColumnIndex;
    }

    public void setStartColumnIndex(int startColumnIndex) {
        this.startColumnIndex = startColumnIndex;
    }

    public int getEndColumnIndex() {
        return endColumnIndex;
    }

    public void setEndColumnIndex(int endColumnIndex) {
        this.endColumnIndex = endColumnIndex;
    }

    public int[] getMergedRegions() {
        return mergedRegions;
    }

    public void setMergedRegions(int[] mergedRegions) {
        this.mergedRegions = mergedRegions;
    }

    public boolean isToMergedRegion() {
        return toMergedRegion;
    }

    public void setToMergedRegion(boolean toMergedRegion) {
        this.toMergedRegion = toMergedRegion;
    }

    @Override
    public String toString() {
        return "CellEntity{" +
                "value='" + value + '\'' +
                ", startRowIndex=" + startRowIndex +
                ", endRowIndex=" + endRowIndex +
                ", startColumnIndex=" + startColumnIndex +
                ", endColumnIndex=" + endColumnIndex +
                ", mergedRegions=" + Arrays.toString(mergedRegions) +
                '}';
    }
}
