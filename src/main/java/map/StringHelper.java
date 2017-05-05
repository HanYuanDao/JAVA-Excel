package map;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Desciption:
 * Author: JasonHan.
 * Creation time: 2017/03/27 16:07:00.
 * © Copyright 2013-2017, Node Supply Chain Management.
 */

public class StringHelper {

    /**
     * remove all spaces and enter and tab.
     * @param source
     * @return
     */
    public static String formatStr(String source) {
        String target = null;
        if (null != source) {
            Pattern p = Pattern.compile("\\s*|\t|\r|\n");
            Matcher m = p.matcher(source);
            target = m.replaceAll("");
        }
        return target;
    }
}
