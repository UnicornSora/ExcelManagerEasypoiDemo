package com.caterpillar.exceldatamanager.logic;

import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import com.caterpillar.exceldatamanager.entity.Subledger;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.*;

/**
 * @create 2018-01-19 0:13
 * @desc
 **/
@Slf4j
public class ExcelReaderUtil {
    //excel2003扩展名
    public static final String EXCEL03_EXTENSION = ".xls";
    //excel2007扩展名
    public static final String EXCEL07_EXTENSION = ".xlsx";

    public static List<Map<String, Object>> mapList;

    /**
     * 每获取一条记录，即打印
     * 在flume里每获取一条记录即发送，而不必缓存起来，可以大大减少内存的消耗，这里主要是针对flume读取大数据量excel来说的
     *
     * @param sheetName
     * @param sheetIndex
     * @param curRow
     * @param cellList
     */
    public static void sendRows(String filePath, String sheetName, int sheetIndex, int curRow, List<String> cellList) {
        ExportParams exportParams = new ExportParams();
        exportParams.setSheetName(sheetName);
        exportParams.setType(ExcelType.XSSF);
        Subledger subledger = new Subledger();
        for (int i = 0; i < cellList.size(); i++) {
            switch (i) {
                case 0:
                    subledger.setChineseSubjectCode(cellList.get(i));
                case 1:
                    subledger.setChineseSubjectDescription(cellList.get(i));
                case 2:
                    subledger.setYear(cellList.get(i));
                case 3:
                    subledger.setMonth(cellList.get(i));
                case 4:
                    subledger.setDay(cellList.get(i));
                case 5:
                    subledger.setErpCertificateNumber(cellList.get(i));
                case 6:
                    subledger.setAbstractMsg(cellList.get(i));
                case 7:
                    subledger.setDebit(cellList.get(i));
                case 8:
                    subledger.setLender(cellList.get(i));
                case 9:
                    subledger.setDirection(cellList.get(i));
                case 10:
                    subledger.setBalance(cellList.get(i));
            }
        }
        if (mapList.size() != 0 && sheetIndex - 1 < mapList.size() && mapList.get(sheetIndex - 1) != null) {
            Map<String, Object> sheetMap = mapList.get(sheetIndex - 1);
            List<Subledger> subledgerList = (List<Subledger>) sheetMap.get("data");
            subledgerList.add(subledger);
        } else {
            Map<String, Object> sheetMap = new HashMap();
            sheetMap.put("title", exportParams);
            sheetMap.put("entity", Subledger.class);
            List<Subledger> subledgerList = new ArrayList<>();
            subledgerList.add(subledger);
            sheetMap.put("data", subledgerList);
            mapList.add(sheetIndex - 1, sheetMap);
        }
    }

    public static void readExcel(MultipartFile file) throws Exception {
        int totalRows = 0;
        String fileName = file.getOriginalFilename();
        if (fileName.endsWith(EXCEL03_EXTENSION)) { //处理excel2003文件
            ExcelXlsReader excelXls = new ExcelXlsReader();
            totalRows = excelXls.process(file);
        } else if (fileName.endsWith(EXCEL07_EXTENSION)) {//处理excel2007文件
            ExcelXlsxReaderWithDefaultHandler excelXlsxReader = new ExcelXlsxReaderWithDefaultHandler();
            totalRows = excelXlsxReader.process(file);
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
        }
        log.info("发送的总行数：" + totalRows);
    }

    public static void exportExcel(HttpServletResponse response, String filename) throws Exception {
        OutputStream os = null;
        try {
            response.setContentType("application/force-download"); // 设置下载类型
            response.setHeader("Content-Disposition", "attachment;filename=" + filename); // 设置文件的名称
            os = response.getOutputStream(); // 输出流
            SXSSFWorkbook wb = new SXSSFWorkbook(1000);//内存中保留 1000 条数据，以免内存溢出，其余写入 硬盘
            // 获得该工作区的第一个sheet
            Sheet sheet1 = wb.createSheet("sheet1");
            int excelRow = 0; //标题行
            Row titleRow = (Row) sheet1.createRow(excelRow++);
            // TODO
//            for (int i = 0; i < columnList.size(); i++) {
//                Cell cell = titleRow.createCell(i);
//                cell.setCellValue(columnList.get(i));
//            }
//            for (int m = 0; m < cycleCount; m++) {
//                List<List<String>> eventStrList = this.convertPageModelStrList();
//                if (eventStrList != null && eventStrList.size() > 0) {
//                    for (int i = 0; i < eventStrList.size(); i++) { //明细行
//                        Row contentRow = (Row) sheet1.createRow(excelRow++);
//                        List<String> reParam = (List<String>) eventStrList.get(i);
//                        for (int j = 0; j < reParam.size(); j++) {
//                            Cell cell = contentRow.createCell(j);
//                            cell.setCellValue(reParam.get(j));
//                        }
//                    }
//                }
//            }
            wb.write(os);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        } finally {
            try {
                if (os != null) {
                    os.close();
                }
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            } // 关闭输出流
        }
    }

    public static void copyToTemp(File file, String tmpDir) throws Exception {
        FileInputStream fis = new FileInputStream(file);
        File file1 = new File(tmpDir);
        if (file1.exists()) {
            file1.delete();
        }
        FileOutputStream fos = new FileOutputStream(tmpDir);
        byte[] b = new byte[1024];
        int n = 0;
        while ((n = fis.read(b)) != -1) {
            fos.write(b, 0, n);
        }
        fis.close();
        fos.close();
    }

    public static void main(String[] args) throws Exception {
        String path = "C:\\Users\\eszha\\Downloads\\subledger_L_for_excel_cat_little.xlsx";
        Date startDate = new Date();
        log.info("read start {}", startDate.getTime());
//        ExcelReaderUtil.readExcel(path);
        Date endDate = new Date();
        log.info("read end {}", endDate.getTime());
        log.info("时间相差 {}", (endDate.getTime() - startDate.getTime()) / 1000);
    }
}
