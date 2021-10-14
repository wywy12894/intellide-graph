package cn.edu.pku.sei.intellide.graph.extraction.docx;

import cn.edu.pku.sei.intellide.graph.extraction.KnowledgeExtractor;
import javafx.util.Pair;
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
    private int minGranularity;
    private int[] levels = new int[7];  // title serial number
    private int[] nums = new int[7];    // entity content key-id
    private String tmpKey, tmpVal;      // min-evel title tmp var

    ArrayList<DocxSection> titleEntity;

    @Override
    public boolean isBatchInsert() {
        return true;
    }

    public static void main(String[] args) {
        DocxExtractor test = new DocxExtractor();
        test.setDataDir("D:\\documents\\SoftwareReuse\\KnowledgeGraph\\HWProject\\doc+req");
        test.extraction();
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

            if(fileName.contains("需求分析")) docType = 0;
            else if(fileName.contains("特性设计")) docType = 1;
            else if(fileName.contains("架构")) docType = 2;
            else continue;

            try {
                minGranularity = getMinGranularity(xd);
                initVar();      // initialize dataStructure defined above
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
        titleEntity = new ArrayList<DocxSection>(minGranularity + 1);
        for(int i = 1;i <= minGranularity;i++) levels[i] = 0;
        for(int i = 0;i <= minGranularity;i++) {
            nums[i] = 0;
            titleEntity.add(i, new DocxSection());
        }
    }

    public void infoFill(int styleID, XWPFParagraph para, Map<String, DocxSection> map) {
        levels[styleID]++; nums[styleID] = 0;
        for(int i = 1;i <= minGranularity;i++) {
            if(titleEntity.get(i) != null && !map.containsKey(titleEntity.get(i).title)) {
                map.put(titleEntity.get(i).title, titleEntity.get(i));
                titleEntity.get(i-1).children.add(titleEntity.get(i));
            }
        }
        titleEntity.set(styleID, new DocxSection());
        titleEntity.get(styleID).title = para.getText();
        titleEntity.get(styleID).level = styleID;
        titleEntity.get(styleID).serial = levels[styleID];
        if(styleID < minGranularity) levels[styleID+1] = 0;
    }

    public boolean validText(String text) {
        if(text == null) return false;
        text = text.replaceAll(" ", "");
        if(text.length() == 0) return false;
        return !text.equals("\t") && !text.equals("\r\n");
    }

    public int getMinGranularity(XWPFDocument xd) {
        int minGranu = 0;
        List<XWPFParagraph> paragraphs = xd.getParagraphs();
        Map<String, Integer> map = new HashMap<>();
        for(XWPFParagraph para: paragraphs) {
            String styleID = para.getStyleID();
            String text = para.getText();
            if(styleID == null || text == null) continue;
            if(styleID.contains("Heading") || (styleID.compareTo("0") >= 0 && styleID.compareTo("6") <= 0)) {
                String key = text + "#" + styleID;
                if(!map.isEmpty() && map.containsKey(key)) map.put(key, map.get(key) + 1);
                else map.put(key, 1);
                minGranu = Math.max(minGranu, Integer.parseInt(styleID.substring(styleID.length()-1, styleID.length())));
            }
        }
        // sorted by title-level in ascending order
        List<Map.Entry<String,Integer>> list = new ArrayList<Map.Entry<String,Integer>>(map.entrySet());
        Collections.sort(list, new Comparator<Map.Entry<String,Integer>>() {
            public int compare(Map.Entry<String, Integer> o1, Map.Entry<String, Integer> o2) {
                return o2.getKey().substring(o2.getKey().length()-1, o2.getKey().length()).compareTo(o1.getKey().substring(o1.getKey().length()-1, o1.getKey().length()));
            }
        });
        for(Map.Entry<String, Integer> entry: list) {
            String key = entry.getKey();
            key = key.substring(key.indexOf("#") + 1, key.length());
            int level = Integer.parseInt(key);
            int value = entry.getValue();
            if(level == minGranu && value > 1) {
                minGranu--;
            }
        }
        return Math.min(minGranu, 5);
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
                if(styleID != null && styleID.compareTo("1") >= 0 && styleID.compareTo(minGranularity+"") <= 0) {
                    titleEntity.get(minGranularity).content.put(tmpKey, tmpVal);
                    tmpKey = "";
                    handleParagraph(bodyElementsIterator, ((XWPFParagraph)(tmpElement)), map);
                    break;
                }
                else {
                    tmpVal += (((XWPFParagraph)(tmpElement)).getText() + '\n');
                }
            }
            else if(tmpElement instanceof XWPFTable) {
                titleEntity.get(minGranularity).content.put(tmpKey, tmpVal);
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
            case "Heading4" :
            case "4" : {
                if(minGranularity >= 4) {
                    infoFill(4, para, map);
                    currentLevel = 4;
                    break;
                }
            }
            case "Heading5" :
            case "5" : {
                if(minGranularity >= 5) {
                    infoFill(5, para, map);
                    currentLevel = 5;
                    break;
                }
            }
            default: {
                // non-title: content attribute
                if(currentLevel < minGranularity) {
                    titleEntity.get(currentLevel).content.put(String.valueOf(++nums[currentLevel]), para.getText());
                }
                else if (currentLevel == minGranularity) {
                    if (titleLevel.equals("Heading" + (minGranularity + 1)) || titleLevel.equals((minGranularity+1) + "")) {
                        // min-level title content
                        handleMinTitle(bodyElementsIterator, para, map);
                    }
                    else {
                        // normal text content
                        titleEntity.get(minGranularity).content.put(String.valueOf(++nums[minGranularity]), para.getText());
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
        if(currentLevel >= 0 && currentLevel <= minGranularity) titleEntity.get(currentLevel).table.add(ja);
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
        for(int i = 0;i <= minGranularity;i++) {
            if(titleEntity.get(i) != null && !map.containsKey(titleEntity.get(i).title)) {
                map.put(titleEntity.get(i).title, titleEntity.get(i));
                if(i > 0) titleEntity.get(i-1).children.add(titleEntity.get(i));
            }
        }
        if(!tmpKey.equals("")) titleEntity.get(minGranularity).content.put(tmpKey, tmpVal);
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
