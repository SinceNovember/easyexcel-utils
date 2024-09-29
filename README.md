> 一行代码通过easyexcel读取/下载excel的工具类

主要方法

**读**

```java
read(InputStream inputStream)
read(InputStream inputStream, BiPredicate<Integer, List<Object>> filter)
read(InputStream inputStream, int headerRowNumber)
read(InputStream inputStream, BiPredicate<Integer, List<Object>> filter, int headerRowNumber)
read(InputStream inputStream, List<String> expectHeaderNames, Supplier<? extends X> mismatchException)
read(InputStream inputStream, BiPredicate<Integer, List<Object>> filter, int headerRowNumber, Consumer<List<String>> headerConsumer)
read(InputStream inputStream, BiPredicate<Integer, List<Object>> filter, int headerRowNumber, Consumer<List<String>> headerConsumer, List<String> expectHeaderNames, Supplier<? extends X> mismatchException)
readWithRowIndex(InputStream inputStream)
readWithRowIndex(InputStream inputStream, BiPredicate<Integer, List<Object>> filter)
readWithRowIndex(InputStream inputStream, int headerRowNumber)
readWithRowIndex(InputStream inputStream, BiPredicate<Integer, List<Object>> filter, int headerRowNumber)
 readWithRowIndex(InputStream inputStream, List<String> expectHeaderNames, Supplier<? extends X> mismatchException)
readWithRowIndex(InputStream inputStream, BiPredicate<Integer, List<Object>> filter, int headerRowNumber, Consumer<List<String>> headerConsumer)
readWithRowIndex(InputStream inputStream, BiPredicate<Integer, List<Object>> filter, int headerRowNumber, Consumer<List<String>> headerConsumer, List<String> expectHeaderNames, Supplier<? extends X> mismatchException)
read(InputStream inputStream, Class<T> clz)
read(InputStream inputStream, Class<T> clz, List<String> expectHeaderNames, Supplier<? extends X> mismatchException)
read(InputStream inputStream, Class<T> clz, int headerRowNumber)
read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter)
read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter, int headerRowNumber)
read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter, int headerRowNumber, Consumer<List<String>> headerConsumer)
read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter, int headerRowNumber, Consumer<List<String>> headerConsumer, List<String> expectHeaderNames, Supplier<? extends X> mismatchException)
readWithRowIndex(InputStream inputStream, Class<T> clz)
readWithRowIndex(InputStream inputStream, Class<T> clz, List<String> expectHeaderNames, Supplier<? extends X> mismatchException)
readWithRowIndex(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter)
readWithRowIndex(InputStream inputStream, Class<T> clz, int headerRowNumber)
readWithRowIndex(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter, int headerRowNumber)
readWithRowIndex(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter, int headerRowNumber, Consumer<List<String>> headerConsumer)
readWithRowIndex(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter, int headerRowNumber, Consumer<List<String>> headerConsumer, List<String> expectHeaderNames, Supplier<? extends X> mismatchException)
read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter, BiConsumer<Integer, T> rowConsumer)
read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter, BiConsumer<Integer, T> rowConsumer, Consumer<List<T>> dataListConsumer)
read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter, BiConsumer<Integer, T> rowConsumer, Consumer<List<T>> dataListConsumer, int headerRowNumber, Consumer<List<String>> headerConsumer)
read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter, BiConsumer<Integer, T> rowConsumer, Consumer<List<T>> dataListConsumer, int headerRowNumber, Consumer<List<String>> headerConsumer, List<String> expectHeaderNames, Supplier<? extends X> mismatchException)
read(InputStream inputStream, Class<T> clz, BiPredicate<Integer, T> filter, BiConsumer<Integer, T> rowConsumer, Consumer<List<T>> dataListConsumer, Integer sheetNo, int headerRowNumber, Consumer<List<String>> headerConsumer, List<String> expectHeaderNames, Supplier<? extends X> mismatchException)

```

**写**

```java
write(String fileName, String sheetName, List<T> dataList)
write(String fileName, String sheetName, Class<T> type, List<T> dataList)
write(String fileName, String sheetName, List<T> dataList, List<String> headers)
writeMultipleSheets(String fileName, Map<String, List<Object>> sheetData, Map<String, List<String>> headerMap)
writeMultipleSheets(String fileName, Class<T> type, Map<String, List<T>> sheetData, Map<String, List<String>> headerMap)
writeByTemplate(String filaName, String templatePath, List<T> dataList)
write(String fileName, Class<T> type, Map<String, List<T>> sheetData, Map<String, List<String>> headerMap)
```

