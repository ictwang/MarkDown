import com.spire.doc.*;
import com.spire.doc.collections.ParagraphCollection;
import com.spire.doc.documents.*;
import com.spire.doc.fields.DocPicture;
import com.spire.doc.fields.TableOfContent;
import com.spire.doc.fields.TextRange;
import org.junit.Test;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

public class MergeWord {


    @Test
    public void merge(){

        String filePath0 = "C:\\Users\\abc\\Desktop\\testWord.docx";
        String filePath1 = "C:\\Users\\abc\\Desktop\\merge1.docx";
        Document doc1 = new Document();
        doc1.loadFromFile(filePath0);
        Document doc2 = new Document(filePath1);

        //将第二个文档的段落作为新的段落添加到第一个文档的最后一个section
        Section lastSection = doc1.getLastSection();
        for(Section section : (Iterable <Section>) doc2.getSections()){
            for(DocumentObject obj :(Iterable<DocumentObject>) section.getBody().getChildObjects()){
                lastSection.getBody().getChildObjects().add(obj.deepClone());
            }
        }
        //在第一个section的第4个paragraph后面加一幅图
        Section sec = doc1.getSections().get(1);
        Paragraph para = sec.getParagraphs().get(4);
        para.appendBreak(BreakType.Line_Break);
        DocPicture picture = para.appendPicture("jietu.png");
        //设置图片宽度
        picture.setWidth(100f);
        //设置图片高度
        picture.setHeight(80f);

        picture.setTitle("这是一朵小花");
        picture.setDistanceLeft(100f);

        doc1.updateTableOfContents();
        //doc1.insertTextFromFile(filePath1 , FileFormat.Docx);
        doc1.saveToFile("result.docx", FileFormat.Docx);
        
    }

    @Test
    public void AddTOC2(){

        //加载已设置大纲级别的测试文档
        Document doc = new Document("merge1.docx");
        Paragraph para = doc.getSections().get(0).getParagraphs().get(0);
        //插入分页符
        para.appendBreak(BreakType.Page_Break);
        //插入分节符 正文部分重新开始计算页码
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
        //toc.applyStyle(String.valueOf(BuiltinStyle.Toc_1));
        doc.updateTableOfContents();

        //定义一个段落样式
        ParagraphStyle paraStyle = new ParagraphStyle(doc);
        paraStyle.setName("myStyle");
        //设置字体及大小
        paraStyle.getCharacterFormat().setFontName("宋体");
        paraStyle.getCharacterFormat().setFontSize(12f);
        //设置字体颜色
        paraStyle.getCharacterFormat().setTextColor(Color.GREEN);
        //设置水平对齐方式
        paraStyle.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Justify);
        //设置行间距
        //paraStyle.getParagraphFormat().setLineSpacing(15.6f);
        doc.getStyles().add(paraStyle);

        Paragraph paragraph = doc.getSections().get(1).getParagraphs().get(3);
        paragraph.applyStyle(paraStyle.getName());

        //查找所有“算法”文本
        TextSelection[] textSelections = doc.findAllString("算法", false, false);
        //设置高亮颜色
        for (TextSelection selection : textSelections) {
            selection.getAsOneRange().getCharacterFormat().setHighlightColor(Color.YELLOW);
        }

        //保存文档
        doc.saveToFile("目录2.docx", FileFormat.Docx_2010);
    }

    @Test
    public void testUtil() throws FileNotFoundException {
        WordUtilImpl u = new WordUtilImpl();

        String filePath0 = "merge1.docx";
        String filePath1 = "dest0806.docx";
        //u.addTocToFormatedWord(filePath0, filePath1);

       /* InputStream inp = new FileInputStream(filePath1);
        int num = u.getPageNum(inp);
        System.out.println("页数：" + num);*/

        /*Document doc = new Document();
        doc.loadFromFile("dest0806.docx");
        System.out.println("页数：" + doc.getPageCountEx());*/

        //u.insertImage("jietu.png", "merge1.docx", "{这是一张图片}","result0806.docx");

        /*String[] srcDocx = {"testWord.docx", "merge1.docx", "result0806.docx"};
        u.mergeDocs(srcDocx, "merge0806.docx");*/

        Document doc = new Document();
        doc.loadFromFile("123.pdf",FileFormat.PDF);
        System.out.println(doc.getPageCountEx() );
    }


}
