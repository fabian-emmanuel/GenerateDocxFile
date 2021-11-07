package com.codewithfibbee.generator.implementation;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.scriptlet4docx.docx.DocxTemplater;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class Test {
    public static void readDocxFile(String fileName, String generatedFileName, Map<String, Object> dataToFill) {
        try {
            File file = new File(fileName);
            FileInputStream fis = new FileInputStream(file.getAbsolutePath());

            XWPFDocument docx = new XWPFDocument(OPCPackage.open(fis));
            XWPFWordExtractor extractor = new XWPFWordExtractor(docx);

            var text = Stream.of(extractor.getText().split("[ \\n\\t\\r.]"))
                    .filter(str -> str.startsWith("$"))
                    .map(x -> x.replaceAll("[^a-zA-Z0-9]", ""))
                    .collect(Collectors.toList());

            Map<String, Object> data = new HashMap<>();

            for(String txt: text){
                for(Map.Entry<String,Object> val: dataToFill.entrySet()) {
                    if (val.getKey().equals(txt))
                        data.put(txt, val.getValue());
                }
            }

            DocxTemplater docxTemplater = new DocxTemplater(new File(fileName));
            docxTemplater.process(new File(generatedFileName), data);


            fis.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    public static void main(String[] args) {
        Map<String, Object> data = new HashMap<>();
        data.put("YourTeamName","test1");
        data.put("description","test2");
        data.put("groupName","test3");
        data.put("thisGroup","test4");
        data.put("howManyPeopleDoYouIdeallyWant","test5");
        data.put("group","test6");
        data.put("background","test6");

        readDocxFile("Template.docx", "generated.docx", data);
    }
}
