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
import java.util.List;

public class DTSExtractor extends KnowledgeExtractor {

    public static final Label DTS = Label.label("DTS");

    public static final String DTS_NO = "dts_no";       // 问题单编号
    public static final String BRIEF_DESC = "brief_desc";       // 简要描述

    public static RelationshipType CREATOR = RelationshipType.withName("creator");
    public static RelationshipType HANDLER = RelationshipType.withName("handler");
    // TODO
    public static RelationshipType SUBMITTER = RelationshipType.withName("submitter");

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
                        JSONObject dts = new JSONObject(s);
                        Node node = this.getDb().createNode();
                        // 建立DTS实体
                        createDTSNode(dts, node);
                        // 建立DTS到Person的链接关系
                        createDTS2PersonRelationship(node, dts.getString("creator"), DTSExtractor.CREATOR);
                        createDTS2PersonRelationship(node, dts.getString("current_handler"), DTSExtractor.HANDLER);
                    }
                    tx.success();
                }
            } catch (JSONException e) {
                e.printStackTrace();
            }
        }
    }

    public void createDTSNode(JSONObject dts, Node node) throws JSONException {
        node.addLabel(DTSExtractor.DTS);
        node.setProperty(DTSExtractor.DTS_NO, dts.getString("dts_no"));
        node.setProperty(DTSExtractor.BRIEF_DESC, dts.getString("brief_desc"));
    }


    public void createDTS2PersonRelationship(Node dtsNode, String personInfo, RelationshipType relationshipType)  {
        if (personInfo.length() == 0)    return;
        // 姓名 工号
        String[] person = personInfo.split("\\s");
        if (person.length != 2 || person[1].length() != 8)     return;
        Node personNode = this.getDb().findNode(PersonExtractor.PERSON, PersonExtractor.ID, person[1]);
        if (personNode == null){
            personNode = this.getDb().createNode();
            personNode.addLabel(PersonExtractor.PERSON);
            personNode.setProperty(PersonExtractor.NAME, person[0]);
            personNode.setProperty(PersonExtractor.ID, person[1]);
        }
        dtsNode.createRelationshipTo(personNode, relationshipType);
    }



}
