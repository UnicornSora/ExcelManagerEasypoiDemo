package com.caterpillar.exceldatamanager.controller;

import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import com.caterpillar.exceldatamanager.entity.Subledger;
import com.caterpillar.exceldatamanager.logic.ExcelLogic;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

@Slf4j
@RestController
public class DataManagerController {

    @RequestMapping("export")
    public void export(HttpServletResponse response) {
        //模拟从数据库获取需要导出的数据
        List<Subledger> subledgerList = new ArrayList<>();
        Subledger subledger = new Subledger();
        subledger.setDirection("dsdsdsd");
        subledgerList.add(subledger);
        //导出操作
        ExcelLogic.exportExcel(subledgerList, null, null, Subledger.class, "测试excel.xls", response);
    }

    @RequestMapping("importExcel")
    public void importExcel(@RequestParam(value = "importExcelFiles", required = false) MultipartFile file, HttpServletResponse response) {
        //解析excel
        log.info("start import {}", file.toString());
        List<Map<String, Object>> subledgerList = ExcelLogic.importExcelMoreSheet(file, 0, 1, Subledger.class);
        //也可以使用MultipartFile,使用 FileUtil.importExcel(MultipartFile file, Integer titleRows, Integer headerRows, Class < T > pojoClass) 导入
        log.info("导入数据一共{}个sheet", subledgerList.size());
        Date date = new Date();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMddHHmmssss");
        String excelExportName = "subledger_L_for_excel_cat_" + dateFormat.format(date) + ".xlsx";
        ExcelLogic.exportExcel(subledgerList, excelExportName, response, ExcelType.XSSF);
        //TODO 保存数据库
    }
}