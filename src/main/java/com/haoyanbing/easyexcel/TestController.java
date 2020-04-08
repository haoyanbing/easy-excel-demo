package com.haoyanbing.easyexcel;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * @author haoyanbing
 * @since 2020/4/8
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
