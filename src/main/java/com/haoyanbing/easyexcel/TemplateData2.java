package com.haoyanbing.easyexcel;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.format.DateTimeFormat;

import java.util.Date;

/**
 * 模板数据
 * @author haoyanbing
 * @since 2020/4/8
 */
public class TemplateData2 {

    @ExcelProperty("ID")
    private Long id;

    @ExcelProperty("用户名")
    private String userName;

    @ExcelProperty("年龄")
    private Integer age;

    @ExcelProperty("性别")
    private Boolean gender;

    @ExcelProperty("生日快乐")
    @DateTimeFormat("yyyy-MM-dd")
    private Date birthDay;

    @ExcelProperty("身高")
    private Float height;

    @ExcelProperty("体重")
    private Double weight;

    @ExcelProperty("地址")
    private String address;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public Boolean getGender() {
        return gender;
    }

    public void setGender(Boolean gender) {
        this.gender = gender;
    }

    public Date getBirthDay() {
        return birthDay;
    }

    public void setBirthDay(Date birthDay) {
        this.birthDay = birthDay;
    }

    public Float getHeight() {
        return height;
    }

    public void setHeight(Float height) {
        this.height = height;
    }

    public Double getWeight() {
        return weight;
    }

    public void setWeight(Double weight) {
        this.weight = weight;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }
}
