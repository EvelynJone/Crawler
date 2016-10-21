package com.zhaoxl.crawler.bean;

/**
 * 类别表
 * Created by zhaoxl on 2016/10/21.
 */
public class Category {

    /**
     * 类型名称
     */
    private String name;
    /**
     * 类型网址
     */
    private String url;
    /**
     * 类型简称
     */
    private String shortName;

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

    public String getShortName() {
        return shortName;
    }

    public void setShortName(String shortName) {
        this.shortName = shortName;
    }
}
