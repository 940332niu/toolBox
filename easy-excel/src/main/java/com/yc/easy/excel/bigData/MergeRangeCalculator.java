package com.yc.easy.excel.bigData;

import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;

/**
 * 合并区间计算工具
 */
public class MergeRangeCalculator {

    public static <T> Map<Integer, List<CellRangeAddress>> calculateMergeRanges(
            List<T> dataList, int startDataRowIndex, int[] mergeColumnIndexes) {

        Map<Integer, List<CellRangeAddress>> mergeMap = new HashMap<>();

        if (dataList == null || dataList.isEmpty()) {
            return mergeMap;
        }

        for (int colIndex : mergeColumnIndexes) {
            List<CellRangeAddress> ranges = new ArrayList<>();

            // 处理只有一行数据的情况
            if (dataList.size() == 1) {
                // 单行数据通常不需要合并，但如果您需要特殊处理可以在这里添加
                continue;
            }

            int startMergeRow = startDataRowIndex;
            String previousValue = getCellValueAsString(dataList.get(0), colIndex);

            for (int i = 1; i < dataList.size(); i++) {
                String currentValue = getCellValueAsString(dataList.get(i), colIndex);
                int currentRow = startDataRowIndex + i;

                if (!Objects.equals(currentValue, previousValue)) {
                    // 值发生变化，记录之前的合并区间
                    if (currentRow - 1 > startMergeRow) { // 只有多行相同才合并
                        ranges.add(new CellRangeAddress(startMergeRow, currentRow - 1, colIndex, colIndex));
                    }
                    startMergeRow = currentRow; // 开始新的区间
                    previousValue = currentValue;
                }

                // 处理最后一行
                if (i == dataList.size() - 1 && Objects.equals(currentValue, previousValue)) {
                    if (currentRow > startMergeRow) {
                        ranges.add(new CellRangeAddress(startMergeRow, currentRow, colIndex, colIndex));
                    }
                }
            }

            if (!ranges.isEmpty()) {
                mergeMap.put(colIndex, ranges);
            }
        }

        return mergeMap;
    }

    private static <T> String getCellValueAsString(T data, int columnIndex) {
        // 这里需要您根据实际的数据结构来获取字段值
        // 如果是Map，使用 ((Map)data).get(key)
        // 如果是POJO，使用反射获取字段值
        // 示例返回：
        return String.valueOf(data); // 实际使用时需要修改
    }
}
