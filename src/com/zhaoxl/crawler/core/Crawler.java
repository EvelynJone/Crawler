package com.zhaoxl.crawler.core;

import com.zhaoxl.crawler.bean.Category;
import com.zhaoxl.crawler.bean.Company;
import com.zhaoxl.crawler.rule.Rule;
import com.zhaoxl.crawler.rule.RuleException;
import com.zhaoxl.crawler.util.TextUtil;
import org.jsoup.Connection;
import org.jsoup.HttpStatusException;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * 爬虫的核心内容
 *  连接url以及抓取页面内容
 * Created by zhaoxl on 2016/10/21.
 */
public class Crawler {


    /**
     * 根据规则连接上url并且获得elements
     * @param rule
     * @return
     */
    public static Elements conn2GetElements(Rule rule) {
        Elements results = new Elements();
        try {
            /**
             * 解析rule
             */
            String url = rule.getUrl();
            String[] params = rule.getParams();
            String[] values = rule.getValues();
            String resultTagName = rule.getResultTagName();
            int type = rule.getType();
            int requestType = rule.getRequestMoethod();

            Connection conn = Jsoup.connect(url);
            // 设置查询参数

            if (params != null)
            {
                for (int i = 0; i < params.length; i++)
                {
                    conn.data(params[i], values[i]);
                }
            }

            // 设置请求类型
            Document doc = null;
            switch (requestType)
            {
                case Rule.GET:
                    doc = conn.timeout(100000).get();
                    break;
                case Rule.POST:
                    doc = conn.timeout(100000).post();
                    break;
            }
            //处理返回数据
            switch (type)
            {
                case Rule.CLASS:
                    results = doc.getElementsByClass(resultTagName);
                    break;
                case Rule.ID:
                    Element result = doc.getElementById(resultTagName);
                    results.add(result);
                    break;
                case Rule.SELECTION:
                    results = doc.select(resultTagName);
                    break;
                default:
                    //当resultTagName为空时默认去body标签
                    if (TextUtil.isEmpty(resultTagName))
                    {
                        results = doc.getElementsByTag("body");
                    }
            }
        }catch (HttpStatusException e) {
            System.out.println("此url不通：" + rule.getUrl());
            e.printStackTrace();
            return null;
        }
        catch (IOException e) {
            e.printStackTrace();
        }
        return results;
    }

    /**
     * 获取公司/类别信息
     * @param rule
     *          rule的规则为：
     *
     * @return
     */
    public static List<Company> crawler(Rule rule) {
        // 进行对rule的必要校验
        validateRule(rule);
        List<Company> companyList =new ArrayList<>();
        Company company = null;
        String  lastPage = rule.getUrl(); // 记录最后一次抓取的信息
        try
        {
            //处理返回数据
            Elements results = conn2GetElements(rule);
            // 获取的数据有可能为null
            if (results == null) {
                return null;
            }
            for (Element result : results)
            {
                //必要的筛选
                String linkHref = result.attr("href");
                String linkText = result.text();

                company = new Company();

                company.setUrl(linkHref);
                company.setName(linkText);

                companyList.add(company);
            }

        } catch (Exception e){
            System.out.println("最后抓取的页面" + lastPage);
            e.printStackTrace();
        }
        return companyList;
    }


    /**
     *
     * @param rule
     * 				rule的规则定于：
     * 						new Rule("url", "table",Rule.SELECTION, Rule.GET)
     * 						"table"和Rule.SELECTION是固定的
     * @param start  开始抓取tr
     * @param end	结束抓取tr
     * @param index	第index个td
     * @return
     */
    public static List<String> extractTable(Rule rule,int start,int end,int index)
    {

        // 进行对rule的必要校验
        validateRule(rule);

        List<String> text = new ArrayList<>();
        String  lastPage = rule.getUrl(); // 记录最后一次抓取的信息
        try
        {
            Elements result = conn2GetElements(rule);
            // 获取的数据有可能为null
            if (result == null) {
                return null;
            }
            Elements links = result.select("tr");
            for (int i = start ; i < end ; i ++ ) {
                Elements tds = links.get(i).select("td");
                text.add(tds.get(index).text());
            }

        }catch (Exception e) {
            System.out.println("最后抓取的页面" + lastPage);
            e.printStackTrace();
        }
        return text;
    }

    /**
     *
     * @param rule
     *          rule中有一个条件是tagName，此方法获取的就是该tagName下的子tag信息
     * @return
     */
    public static List<Category> crawTagnamesChildren(Rule rule)
    {

        // 进行对rule的必要校验
        validateRule(rule);

        List<Category> categoryList =new ArrayList<>();
        Category category = null;
        String  lastPage = rule.getUrl(); // 记录最后一次抓取的信息
        try
        {
            //处理返回数据
            Elements results = conn2GetElements(rule);
            // 获取的数据有可能为null
            if (results == null) {
                return null;
            }
            for (Element result : results)
            {
                // 获取孩子信息
                Element temp= result.child(0);
                //必要的筛选
                String linkHref = temp.attr("href");
                String linkText = temp.text();

                category = new Category();
                category.setUrl(linkHref);
                category.setName(linkText);
                categoryList.add(category);
            }

        } catch (Exception e)
        {
            System.out.println("最后抓取的页面" + lastPage);
            e.printStackTrace();
        }
        return categoryList;
    }

    /**
     * 对传入的参数进行必要的校验
     */
    private static void validateRule(Rule rule)
    {
        String url = rule.getUrl();
        if (TextUtil.isEmpty(url))
        {
            throw new RuleException("url不能为空！");
        }
        if (!url.startsWith("http://"))
        {
            throw new RuleException("url的格式不正确！");
        }

        if (rule.getParams() != null && rule.getValues() != null)
        {
            if (rule.getParams().length != rule.getValues().length)
            {
                throw new RuleException("参数的键值对个数不匹配！");
            }
        }

    }
}
