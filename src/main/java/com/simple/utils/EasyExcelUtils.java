package com.simple.utils;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.read.listener.ReadListener;
import org.apache.commons.compress.utils.Lists;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Predicate;

public class EasyExcelUtils {

    public static List<List<Object>> read(InputStream inputStream) {
        List<List<Object>> dataList = Lists.newArrayList();
        read(inputStream, null, null, (rowIndex, item) -> {
            if (item instanceof LinkedHashMap<?, ?> row) {
                dataList.add(new ArrayList<>(row.values()));
            }
        });
        return dataList;
    }

    public static Map<Integer, List<Object>> readWithRowIndex(InputStream inputStream) {
        Map<Integer, List<Object>> rowIndexToRowMap = new LinkedHashMap<>();
        read(inputStream, null, null, (rowIndex, item) -> {
            if (item instanceof LinkedHashMap<?, ?> row) {
                rowIndexToRowMap.put(rowIndex, new ArrayList<>(row.values()));
            }
        });
        return rowIndexToRowMap;
    }

    public static <T> List<T> read(InputStream inputStream, Class<T> clz) {
        return read(inputStream, clz, null);
    }

    public static <T> List<T> read(InputStream inputStream, Class<T> clz, Predicate<T> predicate) {
        List<T> dataList = Lists.newArrayList();
        read(inputStream, clz, predicate, (rowIndex, item) -> dataList.add(item));
        return dataList;
    }

    public static <T> Map<Integer, T> readWithRowIndex(InputStream inputStream, Class<T> clz, Predicate<T> predicate) {
        Map<Integer, T> rowIndexToRowMap = new LinkedHashMap<>();
        read(inputStream, clz, predicate, rowIndexToRowMap::put);
        return rowIndexToRowMap;
    }

    public static <T> void read(InputStream inputStream, Class<T> clz, Predicate<T> predicate, BiConsumer<Integer, T> biConsumer) {
        EasyExcel.read(inputStream, clz, new ReadListener<T>() {
                    @Override
                    public void invoke(T o, AnalysisContext analysisContext) {
                        if (predicate == null || !predicate.test(o)) {
                            biConsumer.accept(analysisContext.readRowHolder().getRowIndex(), o);
                        }
                    }

                    @Override
                    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                    }
                })
                .sheet()
                .doRead();
    }

}
