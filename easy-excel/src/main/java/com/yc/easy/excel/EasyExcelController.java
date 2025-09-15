package com.yc.easy.excel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.fastjson.JSON;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.SetUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.List;

@RestController
@Slf4j
public class EasyExcelController {

//    public static void main(String[] args) throws UnsupportedEncodingException {
//        String fileName = URLEncoder.encode("aah.xlsx", "UTF-8");
//        List<DownloadData> exportDataList = EasyExcelController.buildExportData();
//        // 设置需要合并的列索引（0-based）
//        log.info(JSON.toJSONString(exportDataList));
//        EasyExcel.write(fileName,DownloadData.class)
//                .registerWriteHandler(new DynamicMergeHandler(true, 1, SetUtils.hashSet(1,2)))
//                .sheet("测试")
//                .doWrite(exportDataList);
//    }

    @PostMapping("/api/easy/excel/export")
    public void getExportFile(HttpServletResponse response) throws IOException {

        // 这里注意 有同学反应使用swagger 会导致各种问题，请直接用浏览器或者用postman
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        // 这里URLEncoder.encode可以防止中文乱码 当然和easyexcel没有关系
        String fileName = URLEncoder.encode("abcd", "UTF-8");
        List<DownloadData> exportDataList = this.buildExportData();
        // 设置需要合并的列索引（0-based）
        log.info(JSON.toJSONString(exportDataList));
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        EasyExcel.write(fileName, DownloadData.class)
                .registerWriteHandler(new DynamicMergeHandler(true, 1, SetUtils.hashSet(1, 2)))
                .sheet("测试")
                .doWrite(exportDataList);
//        EasyExcel.write(response.getOutputStream(), DownloadData.class).sheet("模板").doWrite(data());

    }

    private List<DownloadData> buildExportData() {
        List<DownloadData> downloadDataList = new ArrayList<>();
        //生成的数据  共五条数据 塞入两遍
        //第一条 与后面的数据不重复
        DownloadData data1 = new DownloadData();
        data1.setId("1");
        data1.setFlowName("流程名不重复1");
        data1.setFlowId("LCMCBCF01");
        data1.setUserName("张三");
        data1.setOpTime(new Date());
        downloadDataList.add(data1);

        DownloadData data2 = new DownloadData();
        data2.setId("2");
        data2.setFlowName("流程名重复2条");
        data2.setFlowId("LCMCBCF02");
        data2.setUserName("张三");
        data2.setOpTime(new Date());
        downloadDataList.add(data2);

        DownloadData data3 = new DownloadData();
        data3.setId("3");
        data3.setFlowName("流程名重复2条");
        data3.setFlowId("LCMCBCF02");
        data3.setUserName("李四");
        data3.setOpTime(DateUtils.addDays(data2.getOpTime(), 1));
        downloadDataList.add(data3);

        DownloadData data4 = new DownloadData();
        data4.setId("4");
        data4.setFlowName("流程名重复3条");
        data4.setFlowId("LCMCBCF03");
        data4.setUserName("张三");
        data4.setOpTime(new Date());
        downloadDataList.add(data4);

        DownloadData data5 = new DownloadData();
        data5.setId("5");
        data5.setFlowName("流程名重复3条");
        data5.setFlowId("LCMCBCF03");
        data5.setUserName("李四");
        data5.setOpTime(DateUtils.addDays(data4.getOpTime(), 1));
        downloadDataList.add(data5);

        DownloadData data6 = new DownloadData();
        data6.setId("6");
        data6.setFlowName("流程名重复3条");
        data6.setFlowId("LCMCBCF03");
        data6.setUserName("王五");
        data6.setOpTime(DateUtils.addDays(data5.getOpTime(), 1));
        downloadDataList.add(data6);

        DownloadData data7 = new DownloadData();
        data7.setId("7");
        data7.setFlowName("流程名不重复1");
        data7.setFlowId("LCMCBCF04");
        data7.setUserName("张三");
        data7.setOpTime(new Date());
        downloadDataList.add(data7);
        return downloadDataList;
    }


}
