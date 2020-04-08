# Easyexcel工具类使用说明
[官方github](https://github.com/alibaba/easyexcel)

[官方文档](https://www.yuque.com/easyexcel/doc)
## 读Excel
### 同步读
```java
/**
 * 同步读取Excel数据(适合小数据量的情况)
 * @param inputStream 输入流
 * @param clazz 模板类类型
 * @param <T> 模板类
 * @return 数据
 */
public static <T> List<T> readSync(InputStream inputStream, Class<T> clazz);

/**
 * 同步读取Excel数据(适合小数据量的情况)
 * @param file 文件
 * @param clazz 模板类类型
 * @param <T> 模板类
 * @return 数据
 */
public static <T> List<T> readSync(File file, Class<T> clazz);
```
### 异步读
```java
/**
 * 异步的读Excel数据
 * @param inputStream 输入流
 * @param clazz 模板类类型
 * @param analysisEventListener 解析事件监听器
 * @param <T> 模板类
 */
public static <T> void read(InputStream inputStream, Class<T> clazz, AnalysisEventListener<T> analysisEventListener);

/**
 * 异步的读Excel数据
 * @param file 文件
 * @param clazz 模板类类型
 * @param analysisEventListener 解析事件监听器
 * @param <T> 模板类
 */
public static <T> void read(File file, Class<T> clazz, AnalysisEventListener<T> analysisEventListener);
```
#### AnalysisEventListener 解析事件监听器
```java
/**
 * 解析事件监听器
 */
public class TemplateDataListener extends AnalysisEventListener<TemplateData> {

    /**
     * 每读取300条数据执行一次存储数据库，然后清理list
     */
    private static final int BATCH_COUNT = 300;

    /**
     * 数据临时存储
     */
    private List<TemplateData> list = new ArrayList<>();

    /**
     * 每读取一条数据会调用该方法
     * @param data 数据
     * @param context 上下文
     */
    @Override
    public void invoke(TemplateData data, AnalysisContext context) {
        System.out.println("读取到1条数据");
        list.add(data);
        if (list.size() >= BATCH_COUNT) {
            // 保存数据到数据库
            list.clear();
        }
    }

    /**
     * 数据全部读取完成会调用该方法
     * @param context 上下文
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        if (list.size() > 0) {
            // 保存数据到数据库
            list.clear();
        }
        System.out.println("数据读取完成");
    }
    
    /**
     * 在转换异常 获取其他异常下会调用本接口。抛出异常则停止读取。如果这里不抛出异常则 继续读取下一行。
     * @param exception 异常信息
     * @param context 上下文
     * @throws Exception 异常
     */
    @Override
    public void onException(Exception exception, AnalysisContext context) throws Exception {
        System.out.print("读取数据发生异常:");
        if (exception instanceof ExcelDataConvertException) {
            ExcelDataConvertException excelDataConvertException = (ExcelDataConvertException) exception;
            System.out.println(String.format("第%s行，第%s列解析异常，异常数据:%s", excelDataConvertException.getRowIndex(),
                    excelDataConvertException.getColumnIndex(), excelDataConvertException.getCellData()));
        }
    }
}
```
#### 模板类
```java
public class TemplateData {

    @ExcelProperty("ID")
    private Long id;

    @ExcelProperty("用户名")
    private String userName;

    @ExcelProperty("年龄")
    private Integer age;

    @ExcelProperty("性别")
    private Boolean gender;

    @ExcelProperty("生日")
    @DateTimeFormat("yyyy-MM-dd")
    private Date birthDay;

    @ExcelProperty("身高")
    private Float height;

    @ExcelProperty("体重")
    private Double weight;

    @ExcelProperty("地址")
    private String address;

    @ExcelProperty("分数")
    private BigDecimal score;
    
    // 省略getter和setter方法
    // ......
}
```

## 写Excel
### 单Sheet
```java
/**
 * 写到文件
 * @param file 文件
 * @param clazz 模板类类型
 * @param data 数据
 * @param sheetName sheet名
 * @param <T> 模板类
 */
public static <T> void write(File file, Class<T> clazz, List<T> data, String sheetName);

/**
 * 写到Response
 * @param response response对象
 * @param clazz 模板类类型
 * @param data 数据
 * @param sheetName sheet名
 * @param fileName 文件名
 * @param <T> 模板类
 * @throws IOException io异常
 */
public static <T> void write(HttpServletResponse response, Class<T> clazz, List<T> data, String sheetName, String fileName) throws IOException;
```
### 多Sheet
```java
/**
 * 写到文件-多Sheet
 * @param file 文件
 * @param classes 模板类类型
 * @param data 数据
 * @param sheetNames sheet名
 */
public static void write(File file, List<Class<?>> classes, List<List<?>> data, List<String> sheetNames);

/**
 * 写到Response-多Sheet
 * @param response response对象
 * @param classes 模板类类型
 * @param data 数据
 * @param sheetNames sheet名
 * @param fileName 文件名
 * @throws IOException io异常
 */
public static void write(HttpServletResponse response, List<Class> classes, List<List<?>> data, List<String> sheetNames, String fileName) throws IOException;
```
## 使用示例
```java
/**
 * 使用示例
 */
@RestController
public class TestController {

    /**
     * 异步的读取Excel数据文件
     * @param file 文件
     * @throws IOException io异常
     */
    @PostMapping("read")
    public void read(MultipartFile file) throws IOException {
        EasyExcelUtil.read(file.getInputStream(), TemplateData.class, new TemplateDataListener());
    }

    /**
     * 同步的读取Excel数据文件
     * @param file 文件
     * @return 读取的数据
     * @throws IOException io异常
     */
    @PostMapping("read/sync")
    public List<TemplateData> readSync(@RequestParam("file") MultipartFile file) throws IOException {
        return EasyExcelUtil.readSync(file.getInputStream(), TemplateData.class);
    }

    /**
     * 导出Excel文件
     * @param response response
     * @throws IOException io异常
     */
    @GetMapping("export")
    public void write(HttpServletResponse response) throws IOException {
        EasyExcelUtil.write(response, TemplateData.class, data(), "模板数据", "导出数据");
    }

    /**
     * 导出Excel文件-多Sheet
     * @param response response
     * @throws IOException io异常
     */
    @GetMapping("export/sheet")
    public void write2(HttpServletResponse response) throws IOException {
        EasyExcelUtil.write(response, Arrays.asList(TemplateData.class, TemplateData2.class), Arrays.asList(data(), data2()), Arrays.asList("sheet1", "sheet2"), "导出数据2");
    }

    /**
     * 模拟生成TemplateData数据
     * @return 数据
     */
    private List<TemplateData> data() {
        List<TemplateData> list = new ArrayList<>();
        for (int i = 0; i < 67; i++) {
            TemplateData templateData = new TemplateData();
            templateData.setId((long) i);
            templateData.setUserName("张" + i);
            templateData.setAge(28);
            templateData.setGender(true);
            templateData.setBirthDay(new Date());
            templateData.setHeight(178.8F);
            templateData.setWeight(68.9);
            templateData.setAddress("上海市闵行区虹梅南路833号国民商务花园");
            templateData.setScore(new BigDecimal(1.222));
            list.add(templateData);
        }
        return list;
    }

    /**
     * 模拟生成TemplateData2数据
     * @return 数据
     */
    private List<TemplateData2> data2() {
        List<TemplateData2> list = new ArrayList<>();
        for (int i = 0; i < 67; i++) {
            TemplateData2 templateData = new TemplateData2();
            templateData.setId((long) i);
            templateData.setUserName("张" + i);
            templateData.setAge(28);
            templateData.setGender(true);
            templateData.setBirthDay(new Date());
            templateData.setHeight(178.8F);
            templateData.setWeight(68.9);
            templateData.setAddress("上海市闵行区虹梅南路833号国民商务花园");
            list.add(templateData);
        }
        return list;
    }
}
```