package com.yc.easy.excel.bigData;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.SheetWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.CollectionUtils;

import java.util.List;
import java.util.Map;


@Slf4j
public class CellLevelMergeStrategy implements CellWriteHandler {

    private final Map<Integer, List<CellRangeAddress>> mergeMap;
    private boolean merged = false;

    public CellLevelMergeStrategy(Map<Integer, List<CellRangeAddress>> mergeMap) {
        this.mergeMap = mergeMap;
    }

    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<WriteCellData<?>> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {

        // 只在数据行处理，跳过表头
        if (isHead || merged) {
            return;
        }

        // 在所有单元格写入完成后执行合并
        Sheet sheet = writeSheetHolder.getSheet();
        int lastRowIndex = sheet.getLastRowNum();

        // 检查是否已经写入了所有数据行
        if (lastRowIndex >= writeSheetHolder.getExcelWriteHeadProperty().getHeadMap().size() + cellDataList.size() - 1) {
            performMerge(sheet);
            merged = true; // 标记已合并，避免重复执行
        }
    }

    private void performMerge(Sheet sheet) {
        if (!mergeMap.isEmpty()) {
            mergeMap.forEach((colIndex, ranges) -> {
                if (ranges != null && !ranges.isEmpty()) {
                    ranges.forEach(range -> {
                        try {
                            sheet.addMergedRegionUnsafe(range);
                        } catch (Exception e) {
                            log.warn("合并单元格失败: {}", range.toString(), e);
                        }
                    });
                }
            });
            log.info("单元格合并完成，共合并 {} 个区域", mergeMap.values().stream().mapToInt(List::size).sum());
        }
    }
}
