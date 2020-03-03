package com.demo;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

public class CopyInstance {

    public static void main( String[] args ) throws Exception {

        XWPFDocument doc = new XWPFDocument();
        createHeader( doc, "a公司", "aaaa" );
    }

    public static void createHeader( XWPFDocument doc, String orgFullName, String logoFilePath ) throws Exception {
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy( doc, sectPr );

        XWPFHeader header = headerFooterPolicy.createHeader( headerFooterPolicy.DEFAULT );

        XWPFParagraph paragraph = header.createParagraph();
        paragraph.setAlignment( ParagraphAlignment.LEFT );
        paragraph.setBorderBottom( Borders.THICK );
        XWPFRun run = paragraph.createRun();

        String pic = XWDFClass.class.getClassLoader().getResource("").getPath()+"csdn-logo_.png";
        String imgFile = pic.substring(1);

        File file = new File( imgFile );
        InputStream is = new FileInputStream( file );
        XWPFPicture picture = run.addPicture( is, XWPFDocument.PICTURE_TYPE_JPEG, imgFile, Units.toEMU( 80 ), Units.toEMU( 45 ) );
        String blipID = "";
        for( XWPFPictureData picturedata : header.getAllPackagePictures() ) { // 这段必须有，不然打开的logo图片不显示
            blipID = header.getRelationId( picturedata );
            picture.getCTPicture().getBlipFill().getBlip().setEmbed( blipID );
        }
        run.addTab();
        is.close();

        run.setText( "你好" );
        String path = XWDFClass.class.getClassLoader().getResource("").getPath();

        File file1 = new File(path+"wjh.doc");

        FileOutputStream os = new FileOutputStream(file1);

        doc.write( os );
        os.close();

    }

}
