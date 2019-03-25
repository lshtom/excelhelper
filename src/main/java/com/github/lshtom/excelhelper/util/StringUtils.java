package com.github.lshtom.excelhelper.util;

import java.util.Objects;

/**
 * 字符串工具类
 *
 * @author linshaohao
 * @version 1.0.0
 */
public class StringUtils {

    /**
     * 判断字符串是否为空
     *
     * @param str 字符串
     * @return 布尔结果
     */
    public static boolean isEmpty(String str) {
        if (Objects.isNull(str) || str.equals("")) {
            return true;
        }
        return false;
    }

    /**
     * 判断字符串是否非空
     *
     * @param str 字符串
     * @return 布尔结果
     */
    public static boolean nonEmpty(String str) {
        return !isEmpty(str);
    }
}
