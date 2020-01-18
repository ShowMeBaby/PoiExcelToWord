package com.ExcelToWord.entity;
import com.ExcelToWord.entity.CellEntity;

import java.util.ArrayList;
public class VarMapEntity {
    private String type;
    private String StringValue;
    private ArrayList<ArrayList<CellEntity>> TableValue;

    public VarMapEntity(String type) {
        this.type = type;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getStringValue() {
        return StringValue;
    }

    public void setStringValue(String stringValue) {
        StringValue = stringValue;
    }

    public ArrayList<ArrayList<CellEntity>> getTableValue() {
        return TableValue;
    }

    public void setTableValue(ArrayList<ArrayList<CellEntity>> tableValue) {
        TableValue = tableValue;
    }
}
