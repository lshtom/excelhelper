package com.github.lshtom.excelhelper.constant;

/**
 * Excel文件类型（xls或xlsx）
 *
 * @author linshaohao
 * @version 1.00
 */
public enum ExcelType {

    XLS("0,", "xls"),
    XLSX("1", "xlsx");

    private String code;
    private String suffix;

    ExcelType(String code, String suffix) {
        this.code = code;
        this.suffix = suffix;
    }

    public String getCode() {
        return code;
    }

    public String getSuffix() {
        return suffix;
    }
}
