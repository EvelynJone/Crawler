package com.zhaoxl.crawler.main;

import com.zhaoxl.crawler.bean.Category;
import com.zhaoxl.crawler.bean.Company;
import com.zhaoxl.crawler.core.Crawler;
import com.zhaoxl.crawler.rule.Rule;
import com.zhaoxl.crawler.util.MkdirUtil;
import com.zhaoxl.crawler.util.OutExcelUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

/**
 * Created by zhaoxl on 2016/10/21.
 */
public class Main {
    public static void main(String[] args) {
        long startTime = System.currentTimeMillis();
        // 规则：网址名称为http://www.b2b101.com，抓取class为typelist的属性，以get方式访问
        Rule rule = new Rule("http://www.b2b101.com",
                "typelist",Rule.CLASS, Rule.GET);

        // 获取http://www.b2b101.com网址下所有标签的class为typelist的信息
        List<Category> categoryList = Crawler.crawTagnamesChildren(rule);

        //        List<Company> companyList =null; // 类别下面的所有子页面的跳转地址和名称
        List<Company> results = new ArrayList<>();	// 所有最后需要的结果的集合

        for (Category category : categoryList) {
            MyRunnable thread = new MyRunnable(category);
            thread.run();
            results.addAll(thread.getResults());
        }

        // 导出结果信息，excel导出
        LinkedHashMap<String, String> linkedHashMap = new LinkedHashMap<>();
        linkedHashMap.put("categoryName","行业分类");
        linkedHashMap.put("name","网站名称");
        linkedHashMap.put("url","网站地址");
        linkedHashMap.put("summary","网站简介");
        try {
            XSSFWorkbook wb = OutExcelUtil.getExcel(true, linkedHashMap, results, null);
            String fileName = "综合商贸.xlsx";
            File path = new File(MkdirUtil.getDir() + "/导出Excel");//创建文件夹
            if (!path.exists()) {
                path.mkdirs();//不存在该路径则自动生成
            }
            String path1 = MkdirUtil.getDir() + "/导出Excel" + File.separator + fileName;
            FileOutputStream fileout = new FileOutputStream(path1);
            wb.write(fileout);
            fileout.flush();
            fileout.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        long endTime = System.currentTimeMillis();
        System.out.println("共耗时："+ (endTime - startTime)/1000 + "秒");
    }

    public static void printSingle(Company company) {
        System.out.println(company.getName());
        System.out.println(company.getUrl());
        System.out.println(company.getSummary());
        System.out.println("**************************************");
    }
}
class MyRunnable implements Runnable {

    private Category category;
    private List<Company> results = new ArrayList<>();	// 所有最后需要的结果的集合

    public MyRunnable(Category category) {
        this.category = category;
    }

    public List<Company> getResults() {
        return results;
    }

    public void setResults(List<Company> results) {
        this.results = results;
    }

    @Override
    public void run() {
        List<Company> companyList =null; // 类别下面的所有子页面的跳转地址和名称
        List<Company> pageUrl = null;                // 下一页信息
        boolean isFirstTime = true;                    // 是否是第一次访问，第二次开始的url就要用pageUrl的url值了
        do {
            if (isFirstTime) {
                companyList = Crawler.crawler(new Rule(category.getUrl(), "tit", Rule.CLASS, Rule.GET));
            } else {
                companyList = Crawler.crawler(new Rule(pageUrl.get(0).getUrl(), "tit", Rule.CLASS, Rule.GET));
            }
            for (Company child : companyList) {    // 循环每一个公司的子页面
                // 拿到子页面中table的第1~3位信息，取第二条td信息
                List<String> something = Crawler.extractTable(new Rule(child.getUrl(), "table", Rule.SELECTION, Rule.GET), 0, 3, 1);

                Company temp = new Company();
                if (something != null) {    // 当前请求的url可跑通的情况下，记录该url的table信息
                    temp.setCategoryName(category.getName());
                    temp.setName(something.get(0));
                    temp.setUrl(something.get(1));
                    temp.setSummary(something.get(2));
                } else { // 当前请求的url不可跑通的情况下，记录该url的信息
                    temp.setCategoryName(category.getName());
                    temp.setName(child.getName());
                    temp.setUrl(child.getUrl());
                }

                // 将获取的信息放入结果集合中
                this.results.add(temp);
                Main.printSingle(temp);
            }

            if (isFirstTime) {    // 第一次访问的时候是访问首页内容的下一页
                pageUrl = Crawler.crawler(new Rule(category.getUrl(), "next", Rule.CLASS, Rule.GET));
                isFirstTime = false;
            } else {    // 下一次访问的时候就是首页内容的下一页内容当中的下一页
                pageUrl = Crawler.crawler(new Rule(pageUrl.get(0).getUrl(), "next", Rule.CLASS, Rule.GET));
            }

        } while (pageUrl.size() != 0); // 当下一页的内容为0时，结束循环
    }
}

