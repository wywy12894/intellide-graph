package cn.edu.pku.sei.intellide.graph.extraction.json;

import cn.edu.pku.sei.intellide.graph.extraction.KnowledgeExtractor;
import org.apache.commons.io.FileUtils;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.neo4j.graphdb.*;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class MRExtractor extends KnowledgeExtractor {

    public static final Label MR = Label.label("MR");
    public static final String ID = "id";
    public static final String TITLE = "title";
    public static final String CONTENT = "content";
    public static final String URL = "url";
    public static final RelationshipType REFERENCE = RelationshipType.withName("reference");
    public static final RelationshipType AUTHOR = RelationshipType.withName("author");
    public static final RelationshipType ASSIGNEE = RelationshipType.withName("assignee");

    private static String DTSRegex = "DTS[A-Z0-9]+";
    private static String ARRegex = "AR\\.[A-Za-z0-9\\.]+";

    private static Pattern dtsPattern;
    private static Pattern arPattern;

    static {
        dtsPattern = Pattern.compile(DTSRegex);
        arPattern = Pattern.compile(ARRegex);
    }


    @Override
    public void extraction() {
        for (File file : FileUtils.listFiles(new File(this.getDataDir()), new String[]{"json"}, true)) {
            List<String> jsonContent = new ArrayList<>();
            try {
                jsonContent = FileUtils.readLines(file, "utf-8");
            } catch (IOException e) {
                e.printStackTrace();
            }
            try {
                try(Transaction tx = this.getDb().beginTx()) {
                    for (String s : jsonContent) {
                        JSONObject mr = new JSONObject(s);
                        Node node = this.getDb().createNode();
                        // 建立MR实体
                        createMRNode(mr, node);
                        // 建立MR到DTS/AR的reference关系
                        createMRRelationship(mr.getString("content"), node);
                        // 建立MR到Person的链接关系
                        createMR2PersonRelationship(node, mr.getString("author_id"), MRExtractor.AUTHOR);
                        createMR2PersonRelationship(node, mr.getString("assignee_id"), MRExtractor.ASSIGNEE);
                    }
                    tx.success();
                }

            } catch (JSONException e) {
                e.printStackTrace();
            }
        }
    }

    public void createMRNode(JSONObject mr, Node node) throws JSONException {
        node.addLabel(MRExtractor.MR);
        node.setProperty(MRExtractor.ID, mr.getString("id"));
        node.setProperty(MRExtractor.TITLE, mr.getString("title"));
        node.setProperty(MRExtractor.CONTENT, mr.getString("content"));
        node.setProperty(MRExtractor.URL, mr.getString("merge_request_url"));
    }

    // 建立MR到DTS或AR的关联关系
    public void createMRRelationship(String content, Node mrNode) {

        Set<String> ids = new HashSet<>();
        Matcher dtsMatcher = dtsPattern.matcher(content);
        while(dtsMatcher.find()){
            String dts_no = dtsMatcher.group();
            if (ids.contains(dts_no))   continue;
            ids.add(dts_no);
            //System.out.println(dts_no);
            Node dtsNode = this.getDb().findNode(DTSExtractor.DTS, DTSExtractor.DTS_NO, dts_no);
            if(dtsNode != null) {
                //System.out.println(dtsNode.getProperty("brief_desc"));
                mrNode.createRelationshipTo(dtsNode, MRExtractor.REFERENCE);
            }
        }

        Matcher arMatcher = arPattern.matcher(content);
        while(arMatcher.find()){
            String ar_no = arMatcher.group();
            if (ids.contains(ar_no))   continue;
            ids.add(ar_no);
            Node arNode = this.getDb().findNode(RequirementExtractor.AR, RequirementExtractor.BUSINESS_NO, ar_no);
            if(arNode != null) {
                mrNode.createRelationshipTo(arNode, MRExtractor.REFERENCE);
                //System.out.println(ar_no);
            }
        }
    }

    // TODO: person name
    public void createMR2PersonRelationship(Node mrNode, String personId, RelationshipType relationshipType)  {
        if (personId.length() == 0)    return;
        if (personId.length() == 9){
            personId = personId.substring(1);
        }
        Node personNode = this.getDb().findNode(PersonExtractor.PERSON, PersonExtractor.ID, personId);
        if (personNode == null){
            personNode = this.getDb().createNode();
            personNode.addLabel(PersonExtractor.PERSON);
            personNode.setProperty(PersonExtractor.NAME, "");
            personNode.setProperty(PersonExtractor.ID, personId);
        }
        mrNode.createRelationshipTo(personNode, relationshipType);
    }

}
