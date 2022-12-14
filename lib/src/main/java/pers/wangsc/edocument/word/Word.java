package pers.wangsc.edocument.word;

import pers.wangsc.edocument.util.FileType;

import java.util.List;
import java.util.Map;
import java.util.Objects;

public class Word {
    private WordReader wordReader;
    private WordWriter wordWriter;
    private Boolean valid;

    public Word(String filePath) {
        String fileType = FileType.judge(filePath);
        if (Objects.equals(fileType, "Office 2003")) {
            wordReader = new DocReader(filePath);
        } else if (Objects.equals(fileType, "Office 2007")) {
            wordReader = new DocxReader(filePath);
            wordWriter = new DocxWriter(filePath);
        }
        valid = wordReader != null;
    }

    public List<String> getTableContent(int tableAt, int[] locations) {
        return wordReader.getTableContent(tableAt, locations);
    }

    public void writeParagraph(String text, String fontFamily, Double fontSize) {
        wordWriter.writeParagraph(text, fontFamily, fontSize);
    }

    public void replaceParagraphText(Map<String, String> map) {
        wordWriter.replaceParagraphText(map);
    }

    public Boolean getValid() {
        return valid;
    }

    public void save() {
        wordWriter.save();
    }
}
