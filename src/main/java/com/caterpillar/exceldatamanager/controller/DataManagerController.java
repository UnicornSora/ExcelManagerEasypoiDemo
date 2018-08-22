package com.caterpillar.exceldatamanager.controller;

import com.caterpillar.exceldatamanager.entity.Subledger;
import com.caterpillar.exceldatamanager.logic.ExcelLogic;
import com.caterpillar.exceldatamanager.logic.ExcelReaderUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@Slf4j
@RestController
public class DataManagerController {

    @Autowired
    private ExcelReaderUtil excelReaderUtil;

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
        try {
            Date startDate = new Date();
            log.info("start time {}", startDate.getTime());
            ExcelReaderUtil.mapList = new ArrayList<>();
            Date impotStartDate = new Date();
            log.info("impotStartDate {}", (impotStartDate.getTime() - startDate.getTime()) / 1000);
            excelReaderUtil.readExcel(file);
            Date impotEndDate = new Date();
            log.info("impotEndDate {}", (impotEndDate.getTime() - startDate.getTime()) / 1000);
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMddHHmmssss");
            String excelExportName = "subledger_L_for_excel_cat_" + dateFormat.format(new Date()) + ".xlsx";
            Date exportStartDate = new Date();
            log.info("exportStartDate {}", (exportStartDate.getTime() - startDate.getTime()) / 1000);
            excelReaderUtil.exportExcel(response, excelExportName);
            Date exportEndDate = new Date();
            log.info("exportEndDate {}", (exportEndDate.getTime() - startDate.getTime()) / 1000);
            Date endDate = new Date();
            log.info("end import {}", endDate.getTime());
            log.info("时间相差 {}", (endDate.getTime() - startDate.getTime()) / 1000);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        }
    }
}