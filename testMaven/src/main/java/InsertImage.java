import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.ShapeHorizontalAlignment;
import com.spire.doc.documents.BreakType;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.ParagraphStyle;
import com.spire.doc.fields.DocPicture;

import java.awt.*;

public class InsertImage {

    public static void main(String[] args){
        //创建Document对象
        Document doc = new Document();
        doc.loadFromFile("C:\\Users\\abc\\Desktop\\testWord.docx");
        //添加节
        Section section = doc.addSection();

        Section sec = doc.getSections().get(1);
        Paragraph para = sec.getParagraphs().get(4);
        para.appendBreak(BreakType.Line_Break);   //换行


        //Image.FromFile("C:\\Users\\abc\\Desktop\\jietu.png");
        DocPicture picture = para.appendPicture("C:\\Users\\abc\\Desktop\\jietu.png");
        //设置图片宽度
        picture.setWidth(100f);
        //设置图片高度
        picture.setHeight(80f);
        picture.setHorizontalAlignment(
                ShapeHorizontalAlignment.Center);


      /*  //给第一个段落设置样式
        ParagraphStyle style = new ParagraphStyle(doc);
        style.setName("titleStyle");
        style.getCharacterFormat().setBold(true);
        style.getCharacterFormat().setTextColor(Color.BLUE);
        style.getCharacterFormat().setFontName("Arial");
        style.getCharacterFormat().setFontSize(18f);
        doc.getStyles().add(style);
        paragraph1.applyStyle("titleStyle");

        //给第一个段落和第二个段落设置水平居中对齐方式
        paragraph1.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        paragraph2.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        //设置第一个段落的段后间距
        paragraph1.getFormat().setAfterSpacing(15f);*/

        //保存
        doc.saveToFile("InsertImage.docx", FileFormat.Docx_2013);
    }
}
