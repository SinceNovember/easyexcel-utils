package com.simple.utils;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.BiPredicate;
import java.util.function.Consumer;
import java.util.function.Supplier;
import java.util.stream.Collectors;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.write.metadata.WriteSheet;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import lombok.SneakyThrows;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.compress.utils.Lists;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.util.UriUtils;

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
        List<List<Object>> dataList = Lists.newArrayList();
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
        read(inputStream, null, filter, rowIndexToRowMap::put, null, headerRowNumber, headerConsumer, expectHeaderNames,
            mismatchException);
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
        List<T> dataList = Lists.newArrayList();
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

    public static <T, X extends Throwable> void read(InputStream inputStream,
                                                     Class<T> clz,
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

    public static <T, X extends Throwable> void read(InputStream inputStream,
                                                     Class<T> clz,
                                                     BiPredicate<Integer, T> filter,
                                                     BiConsumer<Integer, T> rowConsumer,
                                                     Consumer<List<T>> dataListConsumer,
                                                     Integer sheetNo,
                                                     int headerRowNumber,
                                                     Consumer<List<String>> headerConsumer,
                                                     List<String> expectHeaderNames,
                                                     Supplier<? extends X> mismatchException) throws X {
        EasyExcel.read(inputStream, clz, new AnalysisEventListener<T>() {
                final List<T> dataList = dataListConsumer == null ? null : Lists.newArrayList();

                @Override
                public void invoke(T object, AnalysisContext analysisContext) {
                    int rowIndex = analysisContext.readRowHolder().getRowIndex() + 1;
                    if (object instanceof LinkedHashMap<?, ?> row) {
                        List<?> values = new ArrayList<>(row.values());
                        if (values.isEmpty() || values.get(0) != null) {
                            @SuppressWarnings("unchecked")
                            T castedObject = (T) values;
                            object = castedObject;
                        }
                    }
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


    // 简化的写入方法（单个sheet，无表头）
    public static <T> void write(String fileName, String sheetName, List<T> dataList) {
        write(fileName, sheetName, null, dataList, null);
    }

    // 简化的写入方法（单个sheet，指定类型）
    public static <T> void write(String fileName, String sheetName, Class<T> type, List<T> dataList) {
        write(fileName, sheetName, type, dataList, null);
    }

    public static <T> void write(String fileName, String sheetName, List<T> dataList, List<String> headers) {
        write(fileName, sheetName, null, dataList, headers);
    }

    // 简化的写入方法（单个sheet，带表头）
    private static <T> void write(String fileName, String sheetName, Class<T> type, List<T> dataList,
                                  List<String> headers) {
        Map<String, List<T>> sheetData = Collections.singletonMap(sheetName, dataList);
        Map<String, List<String>> headerMap = CollectionUtils.isEmpty(headers) ? null :
            Collections.singletonMap(sheetName, headers);
        write(fileName, type, sheetData, headerMap);
    }

    public static <T> void writeMultipleSheets(String fileName, Map<String, List<Object>> sheetData,
                                               Map<String, List<String>> headerMap) {
        write(fileName, null, sheetData, headerMap);
    }

    // 简化的写入方法（多个sheet）
    public static <T> void writeMultipleSheets(String fileName, Class<T> type, Map<String, List<T>> sheetData,
                                               Map<String, List<String>> headerMap) {
        write(fileName, type, sheetData, headerMap);
    }

    public static <T> void writeByTemplate(String filaName, String templatePath, List<T> dataList) {
        HttpServletResponse response =
            ((ServletRequestAttributes) (RequestContextHolder.currentRequestAttributes())).getResponse();
        configureExcelDownloadResponse(response, filaName);
        try (InputStream inputStream = EasyExcelUtils.class.getClassLoader().getResourceAsStream(templatePath);
             ServletOutputStream outputStream = response.getOutputStream()) {
            EasyExcel.write(outputStream)
                .withTemplate(inputStream)
                .sheet()
                .doFill(dataList);
        } catch (Exception e) {
            throw new RuntimeException("write by template error");
        }
    }

    public static <T> void write(String fileName, Class<T> type, Map<String, List<T>> sheetData,
                                 Map<String, List<String>> headerMap) {
        HttpServletResponse response =
            ((ServletRequestAttributes) (RequestContextHolder.currentRequestAttributes())).getResponse();
        configureExcelDownloadResponse(response, fileName);
        try (ExcelWriter writer = EasyExcel.write(response.getOutputStream(), type).build()) {
            int sheetNo = 0;
            for (Map.Entry<String, List<T>> entry : sheetData.entrySet()) {
                String sheetName = entry.getKey();
                List<T> data = entry.getValue();
                List<String> headers = headerMap != null ? headerMap.get(sheetName) : null;

                WriteSheet sheet;
                if (org.apache.commons.collections4.CollectionUtils.isNotEmpty(headers)) {
                    sheet = EasyExcel.writerSheet(sheetNo++, sheetName).head(convertToHead(headers)).build();
                } else {
                    sheet = EasyExcel.writerSheet(sheetNo++, sheetName).build();
                }
                writer.write(data, sheet);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    // 辅助方法：将表头列表转换为EasyExcel需要的格式
    private static List<List<String>> convertToHead(List<String> headers) {
        return headers.stream()
            .map(Collections::singletonList)
            .collect(Collectors.toList());
    }

    private static void configureExcelDownloadResponse(HttpServletResponse response, String fileName) {
        if (response == null) {
            throw new NullPointerException("response is null");
        }
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8");
        response.setHeader("Content-disposition",
            "attachment; filename=" + UriUtils.encode(fileName, "UTF-8"));
        response.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
    }}
