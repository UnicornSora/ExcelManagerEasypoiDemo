package com.caterpillar.exceldatamanager.controller;

import com.caterpillar.exceldatamanager.entity.Subledger;
import com.caterpillar.exceldatamanager.logic.ExcelLogic;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;
import java.util.List;

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
        ExcelLogic.exportExcel(subledgerList, "测试1", "测试sheet1", Subledger.class, "测试excel.xls", response);
    }

    @RequestMapping("importExcel")
    public void importExcel(HttpServletResponse response) {
        String filePath = "D:\\DataFile\\subledger_L_for_excel_cat NQ_201707.xls";
        //解析excel
        List<Subledger> subledgerList = ExcelLogic.importExcel(filePath, 0, 1, Subledger.class);
        //也可以使用MultipartFile,使用 FileUtil.importExcel(MultipartFile file, Integer titleRows, Integer headerRows, Class < T > pojoClass) 导入
        log.info("导入数据一共{}行", subledgerList.size());
        ExcelLogic.exportExcel(subledgerList, "测试1", "测试sheet1", Subledger.class, "测试excel.xls", response);
        //TODO 保存数据库
    }
}