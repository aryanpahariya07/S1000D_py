package com.example;

import org.w3c.dom.*;
import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.*;
import javax.xml.transform.stream.*;
import java.io.*;
import java.util.*;

public class M2 {

    private static final String W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static final String A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static final String REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static String findImageTarget(String imageId, String relsFile) {
        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document doc = builder.parse(new File(relsFile));

            NodeList relationships = doc.getElementsByTagNameNS(REL_NS, "Relationship");
            for (int i = 0; i < relationships.getLength(); i++) {
                Element rel = (Element) relationships.item(i);
                if (rel.getAttribute("Id").equals(imageId)) {
                    return rel.getAttribute("Target");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    public static void processTable(Element tableElement, NodeList trElements, Document doc) {
        Element tgroup = doc.createElement("tgroup");
        Element tbody = doc.createElement("tbody");
        Element thead = doc.createElement("thead");
        int maxColumns = 0;
        boolean firstRow = true;

        for (int i = 0; i < trElements.getLength(); i++) {
            Element tr = (Element) trElements.item(i);
            Element tableRow = doc.createElement("row");

            NodeList tcElements = tr.getElementsByTagNameNS(W_NS, "tc");
            for (int j = 0; j < tcElements.getLength(); j++) {
                Element tc = (Element) tcElements.item(j);
                Element entry = doc.createElement("entry");
                Element para = doc.createElement("para");
                StringBuilder text = new StringBuilder();

                NodeList tElements = tc.getElementsByTagNameNS(W_NS, "t");
                for (int k = 0; k < tElements.getLength(); k++) {
                    Element t = (Element) tElements.item(k);
                    if (t.getTextContent() != null) {
                        text.append(t.getTextContent());
                    }
                }

                para.setTextContent(text.toString().trim());
                entry.appendChild(para);
                tableRow.appendChild(entry);
            }

            if (tableRow.hasChildNodes()) {
                if (firstRow) {
                    thead.appendChild(tableRow);
                    firstRow = false;
                } else {
                    tbody.appendChild(tableRow);
                }
            }

            int tableRowColumns = tableRow.getChildNodes().getLength();
            if (tableRowColumns > maxColumns) {
                maxColumns = tableRowColumns;
            }
        }

        tgroup.appendChild(thead);
        tgroup.appendChild(tbody);
        tgroup.setAttribute("cols", String.valueOf(maxColumns));
        tableElement.appendChild(tgroup);
    }

    public static void main(String[] args) {
        String docxXmlFile = "uploads/word/document.xml";
        String s1000dXmlFile = "s1000d.xml";

        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            DocumentBuilder builder = factory.newDocumentBuilder();
            Document docxDoc = builder.parse(new File(docxXmlFile));
            Document s1000dDoc = builder.newDocument();

            Element s1000dRoot = s1000dDoc.createElement("s1000d");
            s1000dDoc.appendChild(s1000dRoot);

            NodeList bodyElements = docxDoc.getElementsByTagNameNS(W_NS, "body");
            if (bodyElements.getLength() > 0) {
                Element bodyElement = (Element) bodyElements.item(0);

                Element levelledPara1 = null;
                Element levelledPara2 = null;
                String listType = null;

                NodeList children = bodyElement.getChildNodes();
                for (int i = 0; i < children.getLength(); i++) {
                    Node child = children.item(i);
                    if (child.getNodeType() == Node.ELEMENT_NODE) {
                        Element element = (Element) child;

                        if (element.getLocalName().equals("p")) {
                            Element pStyle = (Element) element.getElementsByTagNameNS(W_NS, "pStyle").item(0);
                            Element drawing = (Element) element.getElementsByTagNameNS(W_NS, "drawing").item(0);

                            if (drawing != null) {
                                // Process drawing
                                Element figure = s1000dDoc.createElement("figure");
                                Element graphic = s1000dDoc.createElement("graphic");
                                Element blip = (Element) drawing.getElementsByTagNameNS(A_NS, "blip").item(0);
 
                                if (blip != null && blip.hasAttribute("r:embed")) {
                                    String imageId = blip.getAttribute("r:embed");
                                    String imageRelsFile = "uploads/word/_rels/document.xml.rels";
                                    String imageTarget = findImageTarget(imageId, imageRelsFile);
 
                                    if (imageTarget != null) {
                                        String imageFilePath = "uploads/" + imageTarget;
                                        graphic.setAttribute("infoEntityIdent", imageFilePath);
                                    }
                                }
 
                                figure.appendChild(graphic);
                                if (levelledPara2 != null) {
                                    levelledPara2.appendChild(figure);
                                } else if (levelledPara1 != null) {
                                    levelledPara1.appendChild(figure);
                                }
                            }
                            else if (pStyle != null) {
                                String styleVal = pStyle.getAttribute("w:val");

                                if (styleVal.equals("Heading1")) {
                                    // Process Heading1
                                    if (levelledPara2 != null) {
                                        levelledPara1.appendChild(levelledPara2);
                                    }
                                    if (levelledPara1 != null) {
                                        s1000dRoot.appendChild(levelledPara1);
                                    }

                                    levelledPara1 = s1000dDoc.createElement("LevelledPara");
                                    Element title = s1000dDoc.createElement("title");
                                    title.setTextContent(getTextContent(element));
                                    title.setAttribute("style", "1");
                                    levelledPara1.appendChild(title);
                                } else if (styleVal.equals("Heading2")) {
                                    // Process Heading2
                                    if (levelledPara2 != null) {
                                        levelledPara1.appendChild(levelledPara2);
                                    }

                                    levelledPara2 = s1000dDoc.createElement("LevelledPara");
                                    Element title = s1000dDoc.createElement("title");
                                    title.setTextContent(getTextContent(element));
                                    title.setAttribute("style", "2");
                                    levelledPara2.appendChild(title);
                                } else if (styleVal.equals("Title")) {
                                    // Process Title
                                    Element title = s1000dDoc.createElement("title");
                                    title.setTextContent(getTextContent(element));
                                    s1000dRoot.appendChild(title);
                                }else {
                                    // Process regular paragraph or list item
                                    Element numId = (Element) element.getElementsByTagNameNS(W_NS, "numId").item(0);
                                    
                                    if (numId != null) {
                                        String val = numId.getAttribute("w:val");
                                        if (val.equals("1") || val.equals("2")) {
                                            if (listType == null || !listType.equals(val)) {
                                                Element listElement;
                                                if (val.equals("1")) {
                                                    listType = "1";
                                                    listElement = s1000dDoc.createElement("sequencelist");
                                                } else {
                                                    listType = "2";
                                                    listElement = s1000dDoc.createElement("randomlist");
                                                }
                                                levelledPara1.appendChild(listElement);
                                            }
    
                                            String text = getTextContent(element);
                                            if (!text.isEmpty()) {
                                                Element listitem = s1000dDoc.createElement("listitem");
                                                Element para = s1000dDoc.createElement("para");
                                                para.setTextContent(text);
                                                listitem.appendChild(para);
                                                levelledPara1.getLastChild().appendChild(listitem);
                                            }
                                        }
                                    } else {
                                        String text = getTextContent(element);
                                        if (!text.isEmpty()) {
                                            Element para = s1000dDoc.createElement("para");
                                            para.setTextContent(text);
                                            if (levelledPara2 != null) {
                                                levelledPara2.appendChild(para);
                                            } else if (levelledPara1 != null) {
                                                levelledPara1.appendChild(para);
                                            }
                                        }
                                    }
                                }
                            }
                           
                            
                        } else if (element.getLocalName().equals("tbl")) {
                            // Process table
                            Element tableElement = s1000dDoc.createElement("table");
                            NodeList trElements = element.getElementsByTagNameNS(W_NS, "tr");
                            processTable(tableElement, trElements, s1000dDoc);
                            if (levelledPara2 != null) {
                                levelledPara2.appendChild(tableElement);
                            } else if (levelledPara1 != null) {
                                levelledPara1.appendChild(tableElement);
                            }
                        }
                    }
                }

                // Append remaining elements
                if (levelledPara2 != null) {
                    levelledPara1.appendChild(levelledPara2);
                }
                if (levelledPara1 != null) {
                    s1000dRoot.appendChild(levelledPara1);
                }
            }

            // Write the S1000D XML file
            TransformerFactory transformerFactory = TransformerFactory.newInstance();
            Transformer transformer = transformerFactory.newTransformer();
            transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            transformer.setOutputProperty(OutputKeys.INDENT, "yes");
            transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");

            DOMSource source = new DOMSource(s1000dDoc);
            StreamResult result = new StreamResult(new File(s1000dXmlFile));
            transformer.transform(source, result);

            System.out.println("S1000D XML file created successfully.");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getTextContent(Element element) {
        StringBuilder text = new StringBuilder();
        NodeList tElements = element.getElementsByTagNameNS(W_NS, "t");
        for (int i = 0; i < tElements.getLength(); i++) {
            Element t = (Element) tElements.item(i);
            if (t.getTextContent() != null) {
                text.append(t.getTextContent());
            }
        }
        return text.toString().trim();
    }
}