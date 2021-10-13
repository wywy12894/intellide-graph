package cn.edu.pku.sei.intellide.graph.extraction.docx;

import cn.edu.pku.sei.intellide.graph.extraction.KnowledgeExtractor;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.IBodyElement;

import org.neo4j.graphdb.Label;
import org.neo4j.graphdb.RelationshipType;
import org.neo4j.unsafe.batchinsert.BatchInserter;

import com.alibaba.fastjson.JSONObject;
//import org.json.JSONObject;
import org.json.JSONArray;
import org.json.JSONException;

import javax.print.Doc;
import java.io.FileInputStream;
import java.io.File;
import java.io.IOException;
import java.util.*;

public class DocxExtractor extends KnowledgeExtractor {

    public static final Label RequirementSection = Label.label("RequirementSection");
    public static final Label FeatureSection = Label.label("FeatureSection");
    public static final Label ArchitectureSection = Label.label("ArchitectureSection");
    public static final RelationshipType SUB_DOCX_ELEMENT = RelationshipType.withName("subDocxElement");
    public static final String TITLE = "title";
    public static final String CONTENT = "content";
    public static final String TABLE = "table";
    public static final String LEVEL = "level";
    public static final String SERIAL = "serial";

    /* Auxiliary Data Structures */
    private int docType;
    private int currentLevel;
    private int[] levels = new int[4];  // title serial number
    private int[] nums = new int[5];    // entity content key-id
    private String tmpKey, tmpVal;      // level-4 title tmp var

    ArrayList<DocxSection> titleEntity = new ArrayList<DocxSection>(5);

    @Override
    public boolean isBatchInsert() {
        return true;
    }

    @Override
    public void extraction() {
        for (File file : FileUtils.listFiles(new File(this.getDataDir()), new String[] { "docx" }, true)) {
            String fileName = file.getAbsolutePath().substring(new File(this.getDataDir()).getAbsolutePath().length())
                    .replaceAll("^[/\\\\]+", "");
            
            XWPFDocument xd = null;
            try {
                xd = new XWPFDocument(new FileInputStream(file));
            }
            catch(IOException e) {
                e.printStackTrace();
            }
            Map<String, DocxSection> map = new HashMap<>();
            initVar();      // initialize dataStructure defined above
            if(fileName.contains("需求分析")) docType = 0;
            else if(fileName.contains("特性设计")) docType = 1;
            else if(fileName.contains("架构")) docType = 2;
            else continue;
            try {
                titleEntity.get(0).title = fileName.substring(0, fileName.lastIndexOf("."));
                titleEntity.get(0).level = 0;
                titleEntity.get(0).serial = 0;
                parseDocx(xd, map);
            }
            catch(JSONException e) {
                e.printStackTrace();
            }
            for (DocxSection docxSection : map.values()) {
                if(docxSection.level != -1) docxSection.toNeo4j(this.getInserter());
            }
        }
    }

    public void initVar() {
        currentLevel = 0;
        tmpKey = ""; tmpVal = "";
        titleEntity.clear();
        for(int i = 1;i <= 3;i++) levels[i] = 0;
        for(int i = 0;i <= 3;i++) {
            nums[i] = 0;
            titleEntity.add(i, new DocxSection());
        }
    }

    public void infoFill(int styleID, XWPFParagraph para, Map<String, DocxSection> map) {
        levels[styleID]++; nums[styleID] = 0;
        for(int i = 1;i <= 3;i++) {
            if(titleEntity.get(i) != null && !map.containsKey(titleEntity.get(i).title)) {
                map.put(titleEntity.get(i).title, titleEntity.get(i));
                titleEntity.get(i-1).children.add(titleEntity.get(i));
            }
        }
        titleEntity.set(styleID, new DocxSection());
        titleEntity.get(styleID).title = para.getText();
        titleEntity.get(styleID).level = styleID;
        titleEntity.get(styleID).serial = levels[styleID];
        if(styleID < 3) levels[styleID+1] = 0;
    }

    public boolean validText(String text) {
        if(text == null) return false;
        text = text.replaceAll(" ", "");
        if(text.length() == 0) return false;
        return !text.equals("\t") && !text.equals("\r\n");
    }

    public void handleMinTitle(Iterator<IBodyElement> bodyElementsIterator, XWPFParagraph para, Map<String, DocxSection> map) throws JSONException {
        tmpKey = para.getText();
        if(!validText(tmpKey)) return;
        tmpVal = "";
        IBodyElement tmpElement = null;
        while(bodyElementsIterator.hasNext()) {
            tmpElement = bodyElementsIterator.next();
            if(tmpElement instanceof XWPFParagraph) {
                String styleID = ((XWPFParagraph)(tmpElement)).getStyleID();
                if(styleID != null && (styleID.equals("1") || styleID.equals("2") || styleID.equals("3") || styleID.equals("4"))) {
                    titleEntity.get(3).content.put(tmpKey, tmpVal);
                    tmpKey = "";
                    handleParagraph(bodyElementsIterator, ((XWPFParagraph)(tmpElement)), map);
                    break;
                }
                else {
                    tmpVal += (((XWPFParagraph)(tmpElement)).getText() + '\n');
                }
            }
            else if(tmpElement instanceof XWPFTable) {
                titleEntity.get(3).content.put(tmpKey, tmpVal);
                tmpKey = "";
                handleTable(((XWPFTable)tmpElement));
                break;
            }
        }
    }

    public void handleParagraph(Iterator<IBodyElement> bodyElementsIterator, XWPFParagraph para, Map<String, DocxSection> map) throws JSONException {
        if (!validText(para.getText())) return;

        String titleLevel;
        if (para.getStyleID() == null)
            titleLevel = "-1";
        else
            titleLevel = para.getStyleID();
        switch (titleLevel) {
            case "Heading1NoNumber" :
            case "Heading1" :
            case "1" : {
                infoFill(1, para, map);
                currentLevel = 1;
                break;
            }
            case "Heading2" :
            case "2" : {
                infoFill(2, para, map);
                currentLevel = 2;
                break;
            }
            case "Heading3" :
            case "3" : {
                infoFill(3, para, map);
                currentLevel = 3;
                break;
            }
            default: {
                // non-title: content attribute
                if (currentLevel == 0) {
                    // content between titles0.get(0) and titles0.get(1)
                    titleEntity.get(0).content.put(String.valueOf(++nums[0]), para.getText());
                }
                else if (currentLevel == 1) {
                    titleEntity.get(1).content.put(String.valueOf(++nums[1]), para.getText());
                }
                else if (currentLevel == 2) {
                    titleEntity.get(2).content.put(String.valueOf(++nums[2]), para.getText());
                }
                else if (currentLevel == 3) {
                    if (titleLevel.equals("Heading4") || titleLevel.equals("4")) {
                        // level-4 title content
                        handleMinTitle(bodyElementsIterator, para, map);
                    }
                    else {
                        // normal text content
                        titleEntity.get(3).content.put(String.valueOf(++nums[3]), para.getText());
                    }
                }
            }
        }
    }

    public void handleTable(XWPFTable table) {
        JSONArray ja = new JSONArray();
        String[] lines = table.getText().split("\\r?\\n");
        for(String line : lines) {
            ja.put(line);
        }
        if(currentLevel == 0) titleEntity.get(0).table.add(ja);
        else if(currentLevel == 1) titleEntity.get(1).table.add(ja);
        else if(currentLevel == 2) titleEntity.get(2).table.add(ja);
        else if(currentLevel == 3) titleEntity.get(3).table.add(ja);
    }

    public void parseDocx(XWPFDocument xd, Map<String, DocxSection> map) throws JSONException {
        Iterator<IBodyElement> bodyElementsIterator = xd.getBodyElementsIterator();
        while (bodyElementsIterator.hasNext()) {
            IBodyElement bodyElement = bodyElementsIterator.next();
            if(bodyElement instanceof XWPFTable) {
                handleTable(((XWPFTable) (bodyElement)));
            }
            else if(bodyElement instanceof XWPFParagraph) {
                handleParagraph(bodyElementsIterator, ((XWPFParagraph) (bodyElement)), map);
            }
        }
        for(int i = 0;i <= 3;i++) {
            if(titleEntity.get(i) != null && !map.containsKey(titleEntity.get(i).title)) {
                map.put(titleEntity.get(i).title, titleEntity.get(i));
                if(i > 0) titleEntity.get(i-1).children.add(titleEntity.get(i));
            }
        }
        if(!tmpKey.equals("")) titleEntity.get(3).content.put(tmpKey, tmpVal);
    }

    class DocxSection {
        long node = -1;
        String title = "";
        int level = -1;
        int serial = 0;
        JSONObject content = new JSONObject(new LinkedHashMap<>());
        ArrayList<JSONArray> table = new ArrayList<>();
        ArrayList<DocxSection> children = new ArrayList<>();

        public long toNeo4j(BatchInserter inserter) {
            if(node != -1) return node;
            Map<String, Object> map = new HashMap<>();
            map.put(DocxExtractor.TITLE, title);
            map.put(DocxExtractor.LEVEL, level);
            map.put(DocxExtractor.SERIAL, serial);
            map.put(DocxExtractor.CONTENT, content.toString());
            map.put(DocxExtractor.TABLE, table.toString());
            if(docType == 0) node = inserter.createNode(map, new Label[] { DocxExtractor.RequirementSection });
            else if(docType == 1) node = inserter.createNode(map, new Label[] { DocxExtractor.FeatureSection });
            else if(docType == 2) node = inserter.createNode(map, new Label[] { DocxExtractor.ArchitectureSection });
            for (int i = 0; i < children.size(); i++) {
                DocxSection child = children.get(i);
                if(child.level == -1) continue;
                long childId = child.toNeo4j(inserter);
                Map<String, Object> rMap = new HashMap<>();
                rMap.put(DocxExtractor.SERIAL, i);
                inserter.createRelationship(node, childId, DocxExtractor.SUB_DOCX_ELEMENT, rMap);
            }
            return node;
        }
    }
}
