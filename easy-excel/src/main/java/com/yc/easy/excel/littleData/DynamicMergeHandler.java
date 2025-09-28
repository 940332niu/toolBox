package com.yc.easy.excel.littleData;

import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.merge.AbstractMergeStrategy;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.*;

public class DynamicMergeHandler extends AbstractMergeStrategy {

    /**
     * 相同列合并
     */
    private boolean alikeColumn;

    /**
     * 相同行合并
     */
    private boolean alikeRow;

    /**
     * 开始进行合并的行index
     */
    private int rowIndex;

    /**
     * 跨行合并的列index
     */
    private Set<Integer> columns;

    private int currentRowIndex = 0;

    /**
     * 构造方法,指定合并方式
     *
     * @param alikeColumn
     * @param alikeRow
     * @param rowIndex
     * @param columns
     */
    public DynamicMergeHandler(boolean alikeColumn, boolean alikeRow, int rowIndex, Set<Integer> columns) {
        this.alikeColumn = alikeColumn;
        this.alikeRow = alikeRow;
        this.rowIndex = rowIndex;
        this.columns = columns;
    }

    /**
     * 指定是否进行跨列合并单元格
     *
     * @param alikeColumn
     * @param rowIndex
     */
    public DynamicMergeHandler(boolean alikeColumn, int rowIndex) {
        this(alikeColumn, false, rowIndex, new HashSet<>());
    }

    /**
     * 指定是否进行跨行合并单元格
     *
     * @param alikeRow
     * @param rowIndex
     * @param columns
     */
    public DynamicMergeHandler(boolean alikeRow, int rowIndex, Set<Integer> columns) {
        this(false, alikeRow, rowIndex, columns);
    }


    @Override
    protected void merge(Sheet sheet, Cell cell, Head head, Integer integer) {
        int rowId = cell.getRowIndex();
        currentRowIndex = rowId == currentRowIndex ? currentRowIndex : rowId;
        if (rowIndex > rowId) {
            return;
        }
        int columnId = cell.getColumnIndex();
        // 列合并
        if (alikeColumn && columnId > 0) {
            String currentCellVal = this.getCellVal(cell);
            Cell preCell = cell.getRow().getCell(columnId - 1);
            String preCellVal = this.getCellVal(preCell);
            if (StringUtils.isNotBlank(preCellVal) && StringUtils.isNotBlank(currentCellVal)) {
                // 当前单元格内容与上一个单元格内容相等,进行合并处理
                if (preCellVal.equals(currentCellVal)) {
                    CellRangeAddress rangeAddress = new CellRangeAddress(currentRowIndex, currentRowIndex, columnId - 1, columnId);
                    rangeAddress = this.findExistAddress(sheet, rangeAddress, currentCellVal);
                    if (null != rangeAddress) {
                        sheet.addMergedRegion(rangeAddress);
                    }
                }
            }
        }
        // 行合并
        if (alikeRow && rowIndex < rowId && columns.contains(columnId)) {
            String currentCellVal = this.getCellVal(cell);
            Cell preCell = sheet.getRow(rowId - 1).getCell(columnId);
            String preCellVal = this.getCellVal(preCell);
            if (StringUtils.isNotBlank(preCellVal) && StringUtils.isNotBlank(currentCellVal)) {
                // 当前单元格内容与上一行单元格内容相等,进行合并处理
                if (preCellVal.equals(currentCellVal)) {
                    //sheet.validateMergedRegions();
                    CellRangeAddress rangeAddress = new CellRangeAddress(currentRowIndex - 1, currentRowIndex, columnId, columnId);
                    rangeAddress = this.findExistAddress(sheet, rangeAddress, currentCellVal);
                    if (null != rangeAddress) {
                        sheet.addMergedRegion(rangeAddress);
                    }
                }
            }
        }
    }


    /**
     * 合并单元格地址范围,发现存在相同的地址则进行扩容合并
     *
     * @param sheet
     * @param rangeAddress 单元格合并地址
     * @param currentVal   当前单元格中的值
     * @return
     */
    private CellRangeAddress findExistAddress(Sheet sheet, CellRangeAddress rangeAddress, String currentVal) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        int existIndex = 0;
        Map<Integer, CellRangeAddress> existIdexMap = new LinkedHashMap<>();
        if (null != mergedRegions && !mergedRegions.isEmpty()) {
            //验证当前合并的单元格是否存在重复
            for (CellRangeAddress mergedRegion : mergedRegions) {
                if (mergedRegion.intersects(rangeAddress)) {
                    existIdexMap.put(existIndex, mergedRegion);
                }
                existIndex++;
            }
        }
        if (existIdexMap.isEmpty()) {
            return rangeAddress;
        }
        List<Integer> existIndexList = new ArrayList<>(existIdexMap.size());
        for (Map.Entry<Integer, CellRangeAddress> addressEntry : existIdexMap.entrySet()) {
            CellRangeAddress exist = addressEntry.getValue();
            // 自动进行单元格合并处理
            int firstRow = rangeAddress.getFirstRow();
            int lastRow = rangeAddress.getLastRow();
            int firstColumn = rangeAddress.getFirstColumn();
            int lastColumn = rangeAddress.getLastColumn();

            int firstRow1 = exist.getFirstRow();
            int lastRow1 = exist.getLastRow();
            int firstColumn1 = exist.getFirstColumn();
            int lastColumn1 = exist.getLastColumn();
            // 跨行合并 最后一列相等, 行不相等
            if (lastRow > lastRow1 && lastColumn == lastColumn1) {
                // 检查进行跨行合并的单元格是否已经存在跨列合并
                if (lastColumn > 0 && firstColumn1 != lastColumn1) {
                    // 获取当前单元格的前一列单元格
                    String cellVal = this.getCellVal(sheet.getRow(lastRow).getCell(lastColumn - 1));
                    if (null != cellVal && cellVal.equals(currentVal)) {
                        exist.setLastRow(lastRow);
                    }
                } else {
                    exist.setLastRow(lastRow);
                }
                rangeAddress = exist;
                existIndexList.add(addressEntry.getKey());
            }

            // 跨列合并 行相等,列不相等
            if (lastColumn > lastColumn1 && firstRow == firstRow1) {
                exist.setLastColumn(lastColumn);
                rangeAddress = exist;
                existIndexList.add(addressEntry.getKey());
            }
        }
        // 移除已经存在且冲突的合并数据
        if (existIndexList.isEmpty()) {
            rangeAddress = null;
        } else {
            sheet.removeMergedRegions(existIdexMap.keySet());
        }
        return rangeAddress;
    }

    /**
     * 获取单元格中的内容
     *
     * @param cell
     * @return
     */
    private String getCellVal(Cell cell) {
        String val = null;
        try {
            val = cell.getStringCellValue();
        } catch (Exception e) {
            System.out.printf("读取单元格内容失败:行%d 列%d %n", (cell.getRowIndex() + 1), (cell.getColumnIndex() + 1));
        }
        return val;
    }

}
