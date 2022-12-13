package pers.wangsc.edocument.util;

import java.util.StringTokenizer;

public class StringUtil {
    public static String[] splitByCommonDelimiter(String str) {
        StringBuilder formattedStr = new StringBuilder();
        String delimiter = " ,，/、。|()（）【】[][]";
        StringTokenizer st = new StringTokenizer(str, delimiter);
        while (st.hasMoreTokens()) {
            formattedStr.append(st.nextToken()).append(" ");
        }
        return formattedStr.toString().split(" ");
    }

    public static int[] stringArrayToNoOffsetIntArray(String[] strArr) {
        int[] result = new int[strArr.length];
        for (int i = 0; i < strArr.length; i++) {
            try {
                result[i] = Integer.parseInt(strArr[i]) - 1;
            } catch (NumberFormatException e) {
                e.printStackTrace();
            }
        }
        return result;
    }

    public static int[] delimiterStringToNoOffsetIntArray(String str) {
        String[] strings = StringUtil.splitByCommonDelimiter(str);
        return stringArrayToNoOffsetIntArray(strings);
    }
}