package com.github.lshtom.excelhelper.common;

import com.github.lshtom.excelhelper.constant.ExcelType;
import com.github.lshtom.excelhelper.util.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;

/**
 * Excel元件：Excel所对应的对象实例的抽象
 *
 * @author linshaohao
 * @version 1.0.0
 */
public class ExcelComponent {

    private Workbook wb;
    private Sheet curSheet;
    private Row headerRow;
    private Row dataRow;
    private Cell cell;
    private ExcelType excelType;
    private int headerRowNum;
    private int dataRowNum;

    private static int DEFAULT_HEADER_ROW_NUM = 0;

    /**
     * Excel组件构造器
     *
     * @param excelType Excel类型
     */
    public ExcelComponent(ExcelType excelType) {
        this.excelType = excelType;
        if (excelType == ExcelType.XLS) {
            wb = new HSSFWorkbook();
        } else {
            wb = new XSSFWorkbook();
        }
    }

    /**
     * 获取Workbook对象实例
     *
     * @return
     */
    public Workbook getWorkbook() {
        return wb;
    }

    /**
     * 创建Sheet页
     *
     * @param sheetName Sheet页名称
     * @return Sheet页对象
     */
    public Sheet createSheet(String sheetName) {
        if (StringUtils.nonEmpty(sheetName)) {
            curSheet = wb.createSheet(sheetName);
        } else {
            curSheet = wb.createSheet();
        }
        return curSheet;
    }

    /**
     * 创建Sheet页
     *
     * @return Sheet页对象
     */
    public Sheet createSheet() {
        return createSheet(null);
    }

    /**
     * 获取当前操作的Sheet页对象
     *
     * @return Sheet页对象
     */
    public Sheet getSheet() {
        return curSheet;
    }

    /**
     * 获取Sheet页对象，若无则返回空
     *
     * @param sheetIndex Sheet页下标
     * @return Sheet页对象
     */
    public Sheet getSheet(int sheetIndex) {
        return wb.getSheetAt(sheetIndex);
    }

    /**
     * 获取Sheet页对象，若无则返回空
     *
     * @param sheetName Sheet页名称
     * @return Sheet页对象
     */
    public Sheet getSheet(String sheetName) {
        return wb.getSheet(sheetName);
    }

    /**
     * 重置当前Sheet页对象为指定下标的Sheet页
     * 说明：只有指定下标的Sheet页存在时才会重置成功，否则返回的是原来Sheet页对象
     *
     * @param sheetIndex Sheet页下标
     * @return 重置后的当前Sheet页对象
     */
    public Sheet resetCurSheet(int sheetIndex) {
        Sheet sheet = wb.getSheetAt(sheetIndex);
        if (Objects.nonNull(sheet)) {
            curSheet = sheet;
        }
        return curSheet;
    }

    /**
     * 重置当前Sheet页对象为指定名称的Sheet页
     * 说明：只有指定名称的Sheet页存在时才会重置成功，否则返回的是原来Sheet页对象
     *
     * @param sheetName Sheet页名称
     * @return 重置后的当前Sheet页对象
     */
    public Sheet resetCurSheet(String sheetName) {
        Sheet sheet = wb.getSheet(sheetName);
        if (Objects.nonNull(sheet)) {
            curSheet = sheet;
        }
        return curSheet;
    }

    /**
     * 为当前Sheet页创建表头行
     * 表头行默认为第0行
     *
     * @return Row行对象
     */
    public Row createHeaderRow() {
        return createHeaderRow(curSheet, DEFAULT_HEADER_ROW_NUM);
    }

    /**
     * 为当前Sheet页在指定行号处创建表头行
     *
     * @param headerRowNum 表头行号
     * @return Row行对象
     */
    public Row createHeaderRow(int headerRowNum) {
        this.headerRowNum = headerRowNum;
        return createHeaderRow(curSheet, headerRowNum);
    }

    /**
     * 为指定下标的Sheet页在指定行号处创建表头行
     *
     * @param sheetIndex Sheet页下标
     * @param headerRowNum 表头行号
     * @return Row行对象
     */
    public Row createHeaderRow(int sheetIndex, int headerRowNum) {
        Sheet sheet = getSheet(sheetIndex);
        if (Objects.nonNull(sheet)) {
            return createHeaderRow(sheet, headerRowNum);
        }
        return null;
    }

    /**
     * 为指定名称的Sheet页在指定行号处创建表头行
     *
     * @param sheetName    Sheet页名称
     * @param headerRowNum 表头行号
     * @return Row行对象
     */
    public Row createHeaderRow(String sheetName, int headerRowNum) {
        Sheet sheet = getSheet(sheetName);
        if (Objects.nonNull(sheet)) {
            return createHeaderRow(sheet, headerRowNum);
        }
        return null;
    }

    /**
     * 获取当前Sheet页的表头行
     *
     * @return Row行对象
     */
    public Row getHeaderRow() {
        return headerRow;
    }

    /**
     * 获取指定下标的Sheet页的表头行
     *
     * @param sheetIndex Sheet页下标
     * @param headerRowNum 表头行号
     * @return Row行对象
     */
    public Row getHeaderRow(int sheetIndex, int headerRowNum) {
        Sheet sheet = getSheet(sheetIndex);
        if (Objects.nonNull(sheet)) {
            return sheet.getRow(headerRowNum);
        }
        return null;
    }

    /**
     * 获取指定名称的Sheet页的表头行
     *
     * @param sheetName    Sheet页名称
     * @param headerRowNum 表头行号
     * @return Row行对象
     */
    public Row getHeaderRow(String sheetName, int headerRowNum) {
        Sheet sheet = getSheet(sheetName);
        if (Objects.nonNull(sheet)) {
            return sheet.getRow(headerRowNum);
        }
        return null;
    }

    /**
     * 根据指定的表头行号创建指定Sheet对象的表头行
     *
     * @param sheet        Sheet页对象
     * @param headerRowNum 表头行号
     * @return Row行对象
     */
    private Row createHeaderRow(Sheet sheet, int headerRowNum) {
        this.headerRow = sheet.createRow(headerRowNum);
        this.headerRowNum = headerRowNum;
        return headerRow;
    }

}
