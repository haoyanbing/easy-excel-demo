package com.haoyanbing.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.write.metadata.WriteSheet;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;
import java.util.List;

/**
 * EasyExcel工具类
 * @author haoyanbing
 * @since 2020/4/7
 */
public class EasyExcelUtil {

    /**
     * 同步读取Excel数据(适合小数据量的情况)
     * @param inputStream 输入流
     * @param clazz 模板类类型
     * @param <T> 模板类
     * @return 数据
     */
    public static <T> List<T> readSync(InputStream inputStream, Class<T> clazz) {
        return EasyExcel.read(inputStream).head(clazz).sheet().doReadSync();
    }

    /**
     * 同步读取Excel数据(适合小数据量的情况)
     * @param file 文件
     * @param clazz 模板类类型
     * @param <T> 模板类
     * @return 数据
     */
    public static <T> List<T> readSync(File file, Class<T> clazz) {
        return EasyExcel.read(file).head(clazz).sheet().doReadSync();
    }

    /**
     * 异步的读Excel数据
     * @param inputStream 输入流
     * @param clazz 模板类类型
     * @param analysisEventListener 解析事件监听器
     * @param <T> 模板类
     */
    public static <T> void read(InputStream inputStream, Class<T> clazz, AnalysisEventListener<T> analysisEventListener) {
        EasyExcel.read(inputStream, clazz, analysisEventListener).sheet().doRead();
    }

    /**
     * 异步的读Excel数据
     * @param file 文件
     * @param clazz 模板类类型
     * @param analysisEventListener 解析事件监听器
     * @param <T> 模板类
     */
    public static <T> void read(File file, Class<T> clazz, AnalysisEventListener<T> analysisEventListener) {
        EasyExcel.read(file, clazz, analysisEventListener).sheet().doRead();
    }

    /**
     * 写到文件
     * @param file 文件
     * @param clazz 模板类类型
     * @param data 数据
     * @param sheetName sheet名
     * @param <T> 模板类
     */
    public static <T> void write(File file, Class<T> clazz, List<T> data, String sheetName) {
        EasyExcel.write(file, clazz).sheet(sheetName).doWrite(data);
    }

    /**
     * 写到文件-多Sheet
     * @param file 文件
     * @param classes 模板类类型
     * @param data 数据
     * @param sheetNames sheet名
     */
    public static void write(File file, List<Class<?>> classes, List<List<?>> data, List<String> sheetNames) {
        ExcelWriter excelWriter = EasyExcel.write(file).build();
        for (int i = 0; i < classes.size(); i++) {
            WriteSheet writeSheet = EasyExcel.writerSheet(sheetNames.get(i)).head(classes.get(i)).build();
            excelWriter.write(data.get(i), writeSheet);
        }
        excelWriter.finish();
    }

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
    public static <T> void write(HttpServletResponse response, Class<T> clazz, List<T> data, String sheetName, String fileName) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        fileName = URLEncoder.encode(fileName, "UTF-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        EasyExcel.write(response.getOutputStream(), clazz).sheet(sheetName).doWrite(data);
    }

    /**
     * 写到Response-多Sheet
     * @param response response对象
     * @param classes 模板类类型
     * @param data 数据
     * @param sheetNames sheet名
     * @param fileName 文件名
     * @throws IOException io异常
     */
    public static void write(HttpServletResponse response, List<Class> classes, List<List<?>> data, List<String> sheetNames, String fileName) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        fileName = URLEncoder.encode(fileName, "UTF-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        ExcelWriter excelWriter = EasyExcel.write(response.getOutputStream()).build();
        for (int i = 0; i < classes.size(); i++) {
            WriteSheet writeSheet = EasyExcel.writerSheet(sheetNames.get(i)).head(classes.get(i)).build();
            excelWriter.write(data.get(i), writeSheet);
        }
        excelWriter.finish();
    }

}
