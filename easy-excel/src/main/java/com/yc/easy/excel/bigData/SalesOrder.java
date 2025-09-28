package com.yc.easy.excel.bigData;


import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

/**
 * 销售订单数据模型
 */
@Data
public class SalesOrder {

    @ExcelProperty(value = "订单编号", index = 0)
    private String orderNo;

    @ExcelProperty(value = "客户名称", index = 1)
    private String customerName;

    @ExcelProperty(value = "销售人员", index = 2)
    private String salesPerson;
}
