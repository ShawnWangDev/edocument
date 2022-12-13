package pers.wangsc.edocument.word;

import java.util.List;

public interface WordReader {
    List<String> getTableContent(int tableAt, int[] locations);
}
