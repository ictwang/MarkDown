package Util;

import java.io.InputStream;

public interface WordUtil {

    //插入图片
    public void insertImage(String picUrl, String docUrl, String pos, String destDocUrl);

    //生成目录
    public void addTocToFormatedWord(String docUrl, String destUrl);

    //合并文档
    public void mergeDocs(String[] srcDocxs, String destDocx);

    //获取页码
    public int getPageNum(InputStream fileStream);


}
