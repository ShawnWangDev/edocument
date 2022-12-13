package pers.wangsc.edocument.word;

import java.util.Map;

public interface WordWriter {
    void writeParagraph(String text, String fontFamily, Double fontSize);

    void replaceParagraphText(Map<String, String> map);

    void save();
}
