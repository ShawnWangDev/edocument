package pers.wangsc.edocument;

import org.junit.jupiter.api.Test;
import pers.wangsc.edocument.word.Word;

import java.util.List;

public class WordTest {

    String wordPath = "D:\\_test\\edocument\\test.docx";

    @Test
    void getTableContent() {
        Word word = new Word(wordPath);
        List<String> data = word.getTableContent(0, new int[]{0, 0, 0, 1, 1, 0, 1, 1});
        data.forEach(System.out::println);
    }
}
