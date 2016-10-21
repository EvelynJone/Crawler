package com.zhaoxl.crawler.bean;

/**
 * 公司表
 * Created by zhaoxl on 2016/10/21.
 */
public class Company {

    /**
     * 公司类别名称
     */
    private String categoryName;
    /**
     * 公司名称
     */
    private String name;
    /**
     * 公司网址
     */
    private String url;
    /**
     * 公司简介
     */
    private String summary;

    public String getCategoryName() {
        return categoryName;
    }

    public void setCategoryName(String categoryName) {
        this.categoryName = categoryName;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

    public String getSummary() {
        return summary;
    }

    public void setSummary(String summary) {
        this.summary = summary;
    }
}
