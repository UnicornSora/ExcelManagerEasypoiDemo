package com.caterpillar.exceldatamanager.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class Subledger {

    @Excel(name = "中方科目代码", orderNum = "0")
    private String chineseSubjectCode;

    @Excel(name = "中方科目描述", orderNum = "1")
    private String chineseSubjectDescription;

    @Excel(name = "年", orderNum = "2")
    private String year;

    @Excel(name = "月", orderNum = "3")
    private String month;

    @Excel(name = "日", orderNum = "4")
    private String day;

    @Excel(name = "ERP 凭证号", orderNum = "5")
    private String erpCertificateNumber;

    @Excel(name = "摘             要", orderNum = "6")
    private String abstractMsg;

    @Excel(name = "借方", orderNum = "7")
    private String debit;

    @Excel(name = "贷方", orderNum = "8")
    private String lender;

    @Excel(name = "方向", orderNum = "9")
    private String direction;

    @Excel(name = "余额", orderNum = "10")
    private String balance;

}
