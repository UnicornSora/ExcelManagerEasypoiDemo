package com.caterpillar.exceldatamanager.logic;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.*;

@Slf4j
public class ExcelLogic {

    public static void exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName, boolean isCreateHeader, HttpServletResponse response) {
        ExportParams exportParams = new ExportParams(title, sheetName);
        exportParams.setCreateHeadRows(isCreateHeader);
        defaultExport(list, pojoClass, fileName, response, exportParams);
    }

    public static void exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName, HttpServletResponse response) {
        defaultExport(list, pojoClass, fileName, response, new ExportParams(title, sheetName));
    }

    public static void exportExcel(List<Map<String, Object>> list, String fileName, HttpServletResponse response, ExcelType excelType) {
        if (excelType == null) {
            excelType = ExcelType.HSSF;
        }
        defaultExport(list, fileName, response, excelType);
    }

    private static void defaultExport(List<?> list, Class<?> pojoClass, String fileName, HttpServletResponse response, ExportParams exportParams) {
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, pojoClass, list);
        if (workbook != null) ;
        downLoadExcel(fileName, response, workbook);
    }

    private static void downLoadExcel(String fileName, HttpServletResponse response, Workbook workbook) {
        try {
            response.setCharacterEncoding("UTF-8");
            response.setHeader("content-Type", "application/vnd.ms-excel");
            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            log.error(e.getMessage(), e);
        }
    }

    private static void defaultExport(List<Map<String, Object>> list, String fileName, HttpServletResponse response, ExcelType excelType) {
        Workbook workbook = ExcelExportUtil.exportExcel(list, excelType);
        if (workbook != null) ;
        downLoadExcel(fileName, response, workbook);
    }


    public static <T> List<T> importExcel(String filePath, Integer titleRows, Integer headerRows, Class<T> pojoClass) {
        if (StringUtils.isBlank(filePath)) {
            return null;
        }
        ImportParams params = new ImportParams();
        params.setTitleRows(titleRows);
        params.setHeadRows(headerRows);
        List<T> list = null;
        try {
            list = ExcelImportUtil.importExcel(new File(filePath), pojoClass, params);
        } catch (NoSuchElementException e) {
            log.error("模板不能为空", e);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        }
        return list;
    }

    public static List<Map<String, Object>> importExcelMoreSheet(MultipartFile file, Integer titleRows, Integer headerRows, Class<?> pojoClass) {
        if (file == null) {
            return null;
        }
        List<Map<String, Object>> mapList = new ArrayList<>();
        try {
            Date importTime = new Date();
            log.info("importTime {}", importTime.getTime());
            ImportParams params = new ImportParams();
            log.info("getWorkBook {}", file.toString());
            Workbook xssfWorkbook = getWorkBook(file);
            for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
                log.info("import xssfWorkbook {} {}", numSheet, xssfWorkbook.toString());
                params.setTitleRows(titleRows);
                params.setHeadRows(headerRows);
                params.setStartSheetIndex(numSheet);
                List<Object> importExcel = ExcelImportUtil.importExcelBySax(file.getInputStream(), pojoClass, params);
                Map sheetMap = new HashMap();
                ExportParams exportParams = new ExportParams();
                exportParams.setSheetName(xssfWorkbook.getSheetName(numSheet));
                exportParams.setType(ExcelType.XSSF);
                sheetMap.put("title", exportParams);
                sheetMap.put("entity", pojoClass);
                sheetMap.put("data", importExcel);
                mapList.add(sheetMap);
                Date importOneTime = new Date();
                log.info("import One time {}", importOneTime.getTime());
            }
        } catch (NoSuchElementException e) {
            log.error("模板不能为空", e);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        }
        return mapList;
    }

    /**
     * 得到Workbook对象
     *
     * @param file
     * @return
     * @throws IOException
     */
    public static Workbook getWorkBook(MultipartFile file) throws IOException {
        InputStream is = file.getInputStream();
        Workbook xssfWorkbook = null;
        try {
            xssfWorkbook = new XSSFWorkbook(is);
        } catch (Exception ex) {
            xssfWorkbook = new HSSFWorkbook(is);
        }
        return xssfWorkbook;
    }

    public static <T> List<T> importExcel(MultipartFile file, Integer titleRows, Integer headerRows, Class<T> pojoClass) {
        if (file == null) {
            return null;
        }
        ImportParams params = new ImportParams();
        params.setTitleRows(titleRows);
        params.setHeadRows(headerRows);
        List<T> list = null;
        try {
            list = ExcelImportUtil.importExcel(file.getInputStream(), pojoClass, params);
        } catch (NoSuchElementException e) {
            log.error("excel文件不能为空", e);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        }
        return list;
    }
}
