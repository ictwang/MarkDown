import Util.WordUtil;
import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.documents.*;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.TableOfContent;
import com.spire.doc.fields.TextRange;

import java.awt.*;
import java.io.FileInputStream;
import java.io.InputStream;

public class WordUtilImpl implements WordUtil {

    public void insertImage(String picUrl, String docUrl, String pos, String destDocUrl) {
        //加载word文档
        Document doc = new Document();
        doc.loadFromFile(docUrl);

        //查找文档中的字符串
        TextSelection[] selections = doc.findAllString(pos,true, true);

        //用图片替换文字
        int index = 0;
        TextRange range = null;
        for(Object obj : selections){
            TextSelection textSelection = (TextSelection)obj;
            DocPicture pic = new DocPicture(doc);
            pic.loadImage(picUrl);
            //设置图片宽度
            pic.setWidth(100f);
            //设置图片高度
            pic.setHeight(80f);
            range = textSelection.getAsOneRange();
            index = range.getOwnerParagraph().getChildObjects().indexOf(range);
            range.getOwnerParagraph().getChildObjects().insert(index, pic);
            range.getOwnerParagraph().getChildObjects().remove(range);
        }

        //保存
        doc.saveToFile(destDocUrl, FileFormat.Docx);
    }

    public void addTocToFormatedWord(String docUrl, String destUrl) {
        //加载word文档
        Document doc = new Document();
        doc.loadFromFile(docUrl);

        Paragraph para = doc.getSections().get(0).getParagraphs().get(0);
        //插入分页符
        para.appendBreak(BreakType.Page_Break);
        //插入分节符，正文部分重新开始计算页码
        para.insertSectionBreak(SectionBreakType.New_Column);

        //在文档最前面插入一个段落，写入文本并格式化
        Paragraph paraInserted = new Paragraph(doc);
        //获取TextRange对象 设置文字样式
        TextRange tr= paraInserted.appendText("目 录");
        tr.getCharacterFormat().setBold(true);
        tr.getCharacterFormat().setTextColor(Color.RED);
        tr.getCharacterFormat().setFontName("华文行楷");
        tr.getCharacterFormat().setFontSize(28);
        doc.getSections().get(0).getParagraphs().insert(0,paraInserted);
        paraInserted.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        //通过域代码添加目录表
        TableOfContent toc = new TableOfContent(doc, "{\\o \"1-3\" \\h \\z \\u}");
        doc.getSections().get(0).getParagraphs().get(0).appendTOC(1,3);

        doc.updateTableOfContents();
        doc.saveToFile(destUrl, FileFormat.Docx);
    }

    public void mergeDocs(String[] srcDocxs, String destDocx) {

        int length = null == srcDocxs ? 0 : srcDocxs.length;
        Document doc = new Document(srcDocxs[0]);

        for(int i=1; i<length; i++){
            String docUrl = srcDocxs[i];
            doc.insertTextFromFile(docUrl, FileFormat.Docx);
        }

        doc.saveToFile(destDocx, FileFormat.Docx);
    }

    public int getPageNum(InputStream fileStream) {
        Document doc = new Document();
        //doc.loadFromStream(fileStream, FileFormat.Html);
        doc.loadFromStream(fileStream, FileFormat.Docx);

        return doc.getPageCountEx();
    }
}
