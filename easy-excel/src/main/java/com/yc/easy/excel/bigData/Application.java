package com.yc.easy.excel.bigData;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.util.DateUtils;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class Application {
    public static void main(String[] args) {
        // 1. 准备数据
        List<SalesOrder> dataList = getYourData(); // 您的数据，即使是1条或几条

        // 2. 预计算合并区间
        int startDataRowIndex = 1; // 数据起始行（第0行通常是表头）
        int[] mergeColumns = {0, 1,2}; // 要合并的列索引

        Map<Integer, List<CellRangeAddress>> mergeRanges = MergeRangeCalculator.calculateMergeRanges(dataList, startDataRowIndex, mergeColumns);

        // 3. 执行导出
        EasyExcel.write(DateUtils.format(new Date(),DateUtils.DATE_FORMAT_14) +".xlsx", SalesOrder.class).registerWriteHandler(new CellLevelMergeStrategy(mergeRanges)).sheet("Sheet1").doWrite(dataList);
    }

    private static List<SalesOrder> getYourData() {

        List<SalesOrder> dataList = new ArrayList<SalesOrder>();
        for (int i = 1; i <= 1000; i++) {
            SalesOrder salesOrder = new SalesOrder();
            salesOrder.setOrderNo(String.valueOf(i));
            salesOrder.setSalesPerson("salesPerson" + String.valueOf(i / 1000));
            salesOrder.setCustomerName("customerName" + String.valueOf(i / 1000));
            dataList.add(salesOrder);
//            if (i%3==0){
//                SalesOrder salesOrder1 = new SalesOrder();
//                salesOrder1.setOrderNo(String.valueOf(i ));
//                salesOrder1.setSalesPerson("salesPerson" + String.valueOf(i / 1000));
//                salesOrder1.setCustomerName("customerName" + String.valueOf(i / 1000));
//                dataList.add(salesOrder1);
//            }
        }
        return dataList;
    }
}
