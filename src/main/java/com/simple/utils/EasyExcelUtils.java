package com.simple.utils;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.BiConsumer;
import java.util.function.BiPredicate;
import java.util.function.Consumer;
import java.util.function.Supplier;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import lombok.SneakyThrows;
import org.apache.commons.collections4.CollectionUtils;

public class EasyExcelUtils {
    public static List<List<Object>> read(InputStream inputStream) {
        return read(inputStream, (rowIndex, row) -> true);
    }

    public static List<List<Object>> read(InputStream inputStream, BiPredicate<Integer, List<Object>> filter) {
        return read(inputStream, filter, 1);
    }

    public static List<List<Object>> read(InputStream inputStream, int headerRowNumber) {
        return read(inputStream, (rowIndex, row) -> true, headerRowNumber);
    }


    public static List<List<Object>> read(InputStream inputStream, BiPredicate<Integer, List<Object>> filter,
                                          int headerRowNumber) {
        return read(inputStream, filter, headerRowNumber, null);
    }

    public static <X extends Throwable> List<List<Object>> read(InputStream inputStream,
                                                                List<String> expectHeaderNames,
                                                                Supplier<? extends X> mismatchException) throws X {
        return read(inputStream, null, 1, null, expectHeaderNames, mismatchException);
    }

    public static List<List<Object>> read(InputStream inputStream, BiPredicate<Integer, List<Object>> filter,
                                          int headerRowNumber, Consumer<List<String>> headerConsumer) {
        return read(inputStream, filter, headerRowNumber, headerConsumer, null, null);
    }

    public static <X extends Throwable> List<List<Object>> read(InputStream inputStream,
                                                                BiPredicate<Integer, List<Object>> filter,
                                                                int headerRowNumber,
                                                                Consumer<List<String>> headerConsumer,
                                                                List<String> expectHeaderNames,
                                                                Supplier<? extends X> mismatchException) throws X {
        List<List<Object>> dataList = new ArrayList<>();
        read(inputStream, null, filter, (rowIndex, item) -> {
            if (item instanceof LinkedHashMap<?, ?> row) {
                dataList.add(new ArrayList<>(row.values()));
            }
        }, null, headerRowNumber, headerConsumer, expectHeaderNames, mismatchException);
        return dataList;
    }

    public static Map<Integer, List<Object>> readWithRowIndex(InputStream inputStream) {
        return readWithRowIndex(inputStream, (rowIndex, row) -> true);
    }

    public static Map<Integer, List<Object>> readWithRowIndex(InputStream inputStream,
                                                              BiPredicate<Integer, List<Object>> filter) {
        return readWithRowIndex(inputStream, filter, 1);
    }

    public static Map<Integer, List<Object>> readWithRowIndex(InputStream inputStream,
                                                              int headerRowNumber) {
        return readWithRowIndex(inputStream, (rowIndex, row) -> true, headerRowNumber);
    }


    public static Map<Integer, List<Object>> readWithRowIndex(InputStream inputStream,
                                                              BiPredicate<Integer, List<Object>> filter,
                                                              int headerRowNumber) {
        return readWithRowIndex(inputStream, filter, headerRowNumber, null);
    }

    public static <X extends Throwable> Map<Integer, List<Object>> readWithRowIndex(InputStream inputStream,
                                                                                    List<String> expectHeaderNames,
                                                                                    Supplier<? extends X> mismatchException)
        throws X {
        return readWithRowIndex(inputStream, null, 1, null, expectHeaderNames, mismatchException);
    }

    public static Map<Integer, List<Object>> readWithRowIndex(InputStream inputStream,
                                                              BiPredicate<Integer, List<Object>> filter,
                                                              int headerRowNumber,
                                                              Consumer<List<String>> headerConsumer) {
        return readWithRowIndex(inputStream, filter, headerRowNumber, headerConsumer, null, null);
    }

    public static <X extends Throwable> Map<Integer, List<Object>> readWithRowIndex(InputStream inputStream,
                                                                                    BiPredicate<Integer, List<Object>> filter,
                                                                                    int headerRowNumber,
                                                                                    Consumer<List<String>> headerConsumer,
                                                                                    List<String> expectHeaderNames,
                                                                                    Supplier<? extends X> mismatchException)
        throws X {
        Map<Integer, List<Object>> rowIndexToRowMap = new LinkedHashMap<>();
        read(inputStream, null, filter, (rowIndex, item) -> {
            if (item instanceof LinkedHashMap<?, ?> row) {
                rowIndexToRowMap.put(rowIndex, new ArrayList<>(row.values()));
            }
        }, null, headerRowNumber, headerConsumer, expectHeaderNames, mismatchException);
        return rowIndexToRowMap;
    }

    public static <T> List<T> read(InputStream inputStream, Class<T> clz) {
        return read(inputStream, clz, null);
    }

    public static <T, X extends Throwable> List<T> read(InputStream inputStream, Class<T> clz,
                                                        List<String> expectHeaderNames,
                                                        Supplier<? extends X> mismatchException) throws X {
        return read(inputStream, clz, null, 1, null, expectHeaderNames, mismatchException);
    }

    public static <T> List<T> read(InputStream inputStream, Class<T> clz, int headerRowNumber) {
        return read(inputStream, clz, null, headerRowNumber);
    }

    public static <T> List<T> read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter) {
        return read(inputStream, clz, filter, 1, null, null, null);
    }

    public static <T> List<T> read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter,
                                   int headerRowNumber) {
        return read(inputStream, clz, filter, headerRowNumber, null, null, null);
    }

    public static <T> List<T> read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter,
                                   int headerRowNumber, Consumer<List<String>> headerConsumer) {
        return read(inputStream, clz, filter, headerRowNumber, headerConsumer, null, null);
    }

    public static <T, X extends Throwable> List<T> read(InputStream inputStream, Class<T> clz,
                                                        BiPredicate<Integer, T> filter,
                                                        int headerRowNumber, Consumer<List<String>> headerConsumer,
                                                        List<String> expectHeaderNames,
                                                        Supplier<? extends X> mismatchException) throws X {
        List<T> dataList = new ArrayList<>();
        read(inputStream, clz, filter, (rowIndex, item) -> dataList.add(item), null, headerRowNumber,
            headerConsumer, expectHeaderNames, mismatchException);
        return dataList;
    }

    public static <T> Map<Integer, T> readWithRowIndex(InputStream inputStream, Class<T> clz) {
        return readWithRowIndex(inputStream, clz, null);
    }

    public static <T, X extends Throwable> Map<Integer, T> readWithRowIndex(InputStream inputStream, Class<T> clz,
                                                                            List<String> expectHeaderNames,
                                                                            Supplier<? extends X> mismatchException)
        throws X {
        return readWithRowIndex(inputStream, clz, null, 1, null, expectHeaderNames, mismatchException);
    }

    public static <T> Map<Integer, T> readWithRowIndex(InputStream inputStream, Class<T> clz,
                                                       BiPredicate<Integer, T> filter) {
        return readWithRowIndex(inputStream, clz, filter, 1);
    }

    public static <T> Map<Integer, T> readWithRowIndex(InputStream inputStream, Class<T> clz,
                                                       int headerRowNumber) {
        return readWithRowIndex(inputStream, clz, null, headerRowNumber);
    }

    public static <T> Map<Integer, T> readWithRowIndex(InputStream inputStream, Class<T> clz,
                                                       BiPredicate<Integer, T> filter,
                                                       int headerRowNumber) {
        return readWithRowIndex(inputStream, clz, filter, headerRowNumber, null);
    }

    public static <T> Map<Integer, T> readWithRowIndex(InputStream inputStream, Class<T> clz,
                                                       BiPredicate<Integer, T> filter,
                                                       int headerRowNumber, Consumer<List<String>> headerConsumer) {
        return readWithRowIndex(inputStream, clz, filter, headerRowNumber, headerConsumer, null, null);
    }


    public static <T, X extends Throwable> Map<Integer, T> readWithRowIndex(InputStream inputStream, Class<T> clz,
                                                                            BiPredicate<Integer, T> filter,
                                                                            int headerRowNumber,
                                                                            Consumer<List<String>> headerConsumer,
                                                                            List<String> expectHeaderNames,
                                                                            Supplier<? extends X> mismatchException)
        throws X {
        Map<Integer, T> rowIndexToRowMap = new LinkedHashMap<>();
        read(inputStream, clz, filter, rowIndexToRowMap::put, null, headerRowNumber, headerConsumer, expectHeaderNames,
            mismatchException);
        return rowIndexToRowMap;
    }

    public static <T> void read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter,
                                BiConsumer<Integer, T> rowConsumer) {
        read(inputStream, clz, filter, rowConsumer, null, 1, null);
    }

    public static <T> void read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter,
                                BiConsumer<Integer, T> rowConsumer, Consumer<List<T>> dataListConsumer) {
        read(inputStream, clz, filter, rowConsumer, dataListConsumer, 1, null);
    }

    public static <T> void read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter,
                                BiConsumer<Integer, T> rowConsumer, Consumer<List<T>> dataListConsumer,
                                int headerRowNumber, Consumer<List<String>> headerConsumer) {
        read(inputStream, clz, filter, rowConsumer, dataListConsumer, headerRowNumber, headerConsumer, null, null);
    }

    public static <T, X extends Throwable> void read(InputStream inputStream, Class<T> clz,
                                                     BiPredicate<Integer, T> filter,
                                                     BiConsumer<Integer, T> rowConsumer,
                                                     Consumer<List<T>> dataListConsumer,
                                                     int headerRowNumber,
                                                     Consumer<List<String>> headerConsumer,
                                                     List<String> expectHeaderNames,
                                                     Supplier<? extends X> mismatchException) throws X {
        read(inputStream, clz, filter, rowConsumer, dataListConsumer, null, headerRowNumber, headerConsumer,
            expectHeaderNames, mismatchException);
    }

    @SneakyThrows
    public static <T, X extends Throwable> void read(InputStream inputStream, Class<T> clz,
                                                     BiPredicate<Integer, T> filter,
                                                     BiConsumer<Integer, T> rowConsumer,
                                                     Consumer<List<T>> dataListConsumer,
                                                     Integer sheetNo,
                                                     int headerRowNumber,
                                                     Consumer<List<String>> headerConsumer,
                                                     List<String> expectHeaderNames,
                                                     Supplier<? extends X> mismatchException) throws X {
        EasyExcel.read(inputStream, clz, new AnalysisEventListener<T>() {
                final List<T> dataList = dataListConsumer == null ? null : new ArrayList<>();

                @Override
                public void invoke(T object, AnalysisContext analysisContext) {
                    int rowIndex = analysisContext.readRowHolder().getRowIndex() + 1;
                    if (filter == null || filter.test(rowIndex, object)) {
                        if (rowConsumer != null) {
                            rowConsumer.accept(rowIndex, object);
                        }
                        if (dataListConsumer != null) {
                            dataList.add(object);
                        }
                    }
                }

                @Override
                public void doAfterAllAnalysed(AnalysisContext analysisContext) {
                    if (dataListConsumer != null) {
                        dataListConsumer.accept(dataList);
                    }
                }

                @SneakyThrows
                @Override
                public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
                    if (!Objects.equals(context.readRowHolder().getRowIndex() + 1, headerRowNumber)) {
                        return;
                    }
                    if (headerConsumer != null) {
                        headerConsumer.accept(new ArrayList<>(headMap.values()));
                    }
                    if (CollectionUtils.isNotEmpty(expectHeaderNames)) {
                        if (mismatchException != null &&
                            !CollectionUtils.isEqualCollection(expectHeaderNames, new ArrayList<>(headMap.values()))) {
                            throw mismatchException.get();
                        }
                    }
                }
            })
            .sheet(sheetNo)
            .headRowNumber(headerRowNumber)
            .doRead();
    }
}
