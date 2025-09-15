package com.yc.easy.excel;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

@Data
public class DownloadData {

    @ExcelProperty(value = "序号",index = 0)
    private String id;

    @ExcelProperty(value = "流程名称",index = 1)
    private String flowName;

    @ExcelProperty(value = "流程Id",index = 2)
    private String flowId;

    @ExcelProperty(value = "参与者",index = 3)
    private String userName;

    @ExcelProperty(value = "操作时间",index = 4)
    private Date opTime;
}
