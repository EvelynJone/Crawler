package com.zhaoxl.crawler.util;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Created by zhaoxl on 2016/10/21.
 */
public class MkdirUtil {

    /**
     * 在E盘创建以当前年月日时分秒下的dir
     *
     * @return
     */
    public static String getDir() {
        String path="E:/"+new SimpleDateFormat("yyyy/MM/dd/HH/mm/ss").format(new Date());
        return path;
    }

    public static void main(String[] args) {
        System.out.println(getDir());
    }

}
