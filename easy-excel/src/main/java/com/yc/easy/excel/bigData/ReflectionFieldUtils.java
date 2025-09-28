package com.yc.easy.excel.bigData;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.extern.slf4j.Slf4j;

import java.lang.reflect.Field;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 反射工具类 - 用于获取Excel列对应的字段值
 */
@Slf4j
public class ReflectionFieldUtils {

    private static final Map<Class<?>, List<FieldInfo>> FIELD_CACHE = new ConcurrentHashMap<>();

    private static class FieldInfo {
        private Field field;
        private int columnIndex;

        public FieldInfo(Field field, int columnIndex) {
            this.field = field;
            this.columnIndex = columnIndex;
            this.field.setAccessible(true);
        }

        public Object getValue(Object obj) throws IllegalAccessException {
            return field.get(obj);
        }
    }

    public static Object getCellValue(Object data, int columnIndex) {
        if (data == null) return null;

        try {
            List<FieldInfo> fieldInfos = getFieldInfos(data.getClass());
            for (FieldInfo fieldInfo : fieldInfos) {
                if (fieldInfo.columnIndex == columnIndex) {
                    return fieldInfo.getValue(data);
                }
            }
            return null;
        } catch (Exception e) {
            log.error("获取字段值失败，列索引: {}, 类: {}", columnIndex, data.getClass().getSimpleName(), e);
            return null;
        }
    }

    public static String getCellValueAsString(Object data, int columnIndex) {
        Object value = getCellValue(data, columnIndex);
        return convertToString(value);
    }

    private static String convertToString(Object value) {
        if (value == null) return "";
        if (value instanceof String) return (String) value;
        return value.toString();
    }

    private static List<FieldInfo> getFieldInfos(Class<?> clazz) {
        return FIELD_CACHE.computeIfAbsent(clazz, ReflectionFieldUtils::doGetFieldInfos);
    }

    private static List<FieldInfo> doGetFieldInfos(Class<?> clazz) {
        List<FieldInfo> fieldInfos = new ArrayList<>();
        List<Field> allFields = getAllFields(clazz);

        for (int i = 0; i < allFields.size(); i++) {
            Field field = allFields.get(i);
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            int columnIndex = (excelProperty != null && excelProperty.index() >= 0) ?
                    excelProperty.index() : i;
            fieldInfos.add(new FieldInfo(field, columnIndex));
        }

        fieldInfos.sort(Comparator.comparingInt(f -> f.columnIndex));
        return fieldInfos;
    }

    private static List<Field> getAllFields(Class<?> clazz) {
        List<Field> fields = new ArrayList<>();
        Class<?> currentClass = clazz;
        while (currentClass != null && currentClass != Object.class) {
            Collections.addAll(fields, currentClass.getDeclaredFields());
            currentClass = currentClass.getSuperclass();
        }
        return fields;
    }
}
