package com.task1.xlReading;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;


public class ReadingXLFile {
  public static void main(String args[])
  {
	         ReadingXLFile obj=new ReadingXLFile();
	         List<String> list=new LinkedList<String>(); 
	         list=obj.xlFileReading();
	         System.out.println(list);
	       
	         obj.creatingXMLFile(list);
  }
  
  public List<String> xlFileReading()
  {
	      List<String> xpathList=new LinkedList<String>();
	  try {
	      FileInputStream fis=new FileInputStream(new File("C:\\Users\\apavani\\Desktop\\MISMO 2_6 to MISMO 3_3.xlsx"));
		  XSSFWorkbook workbook=new XSSFWorkbook(fis);
	      XSSFSheet sheet=workbook.getSheetAt(0);
	      DataFormatter formatter = new DataFormatter();
	     
	      for(Row row:sheet)
		  {
			XSSFCell cell=(XSSFCell) row.getCell(6);
			String value = formatter.formatCellValue(cell);
			xpathList.add(value);
			
			  }
	      workbook.close();
		  fis.close();
		  }
	      
	  catch(IOException io)
	  {
		  System.out.println(io);
	  }
	return xpathList;
  }
  
  public  void creatingXMLFile(List<String> list)
  {    
	 
	  try {
		  DocumentBuilder documentBuilder=DocumentBuilderFactory.newInstance().newDocumentBuilder();
	      Document doc=documentBuilder.newDocument();
	      //XPath xpath=XPathFactory.newInstance().newXPath();
	      Iterator<String> it=list.iterator();
	      it.next();
	      int flag=1;
	      restart:
			 while(it.hasNext())
			 {
				 
				 String value=it.next();
				 if(value=="")
				 {//System.out.println(2);
					 continue restart;
				 }
				
				 if(value!="" && flag==1 )
				 {
					 String[] st=value.split("/");
					 
				     System.out.println(st[0]);
				      Element root= doc.createElement(st[0]);
				      doc.appendChild((Node)root);
					  for(int i=1;i<st.length;i++)
					  {  
						  
						   Element element=doc.createElement(st[i]);
						   root.appendChild(element);
						  root=element;
						 
					  }
					  
				 }
				 
				 if(flag>1)
				 {  
					 int i=0;
					Element e=doc.getDocumentElement();
			
					//String s=e.getFirstChild().toString();
					//System.out.println(s);
					//Element e2=(Element) e.getFirstChild();
					//System.out.println(e2.getFirstChild());
					
					 String[] st=value.split("/");
					
					if(st[i].contains("Not Found")==false)
					{
						
					 while(st[i].equals(e.getFirstChild().toString()))
					 {
						 Element e2=(Element) e.getFirstChild();
						 e=e2;
						 i++;
						// System.out.println(2);
						 System.out.println(st[i]+"-----"+e2);
					 }
				   }
				 }
				flag++;
			 }
	      
	    /*  String path="MESSAGE/ABOUT_VERSIONS/ABOUT_VERSION/MISMOVersionIdentifier";
	      String[] st=path.split("/");
	     
	      Element root= doc.createElement(st[0]);
	      doc.appendChild((Node)root);
		  for(int i=1;i<st.length;i++)
		  {  
		    	Element element1=doc.createElement(st[i]);
			   root.appendChild((Node)element1);
			  
			   Element element=doc.createElement(st[i]);
			   root.appendChild(element);
			  root=element;
			 
		  }
	     */
		
			 Transformer transformer=TransformerFactory.newInstance().newTransformer();
			  DOMSource source=new DOMSource(doc);
			  StreamResult Sresult=new StreamResult(new File("D:\\Assignment1\\mismoXML.xml"));
			  transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			  transformer.transform(source, Sresult);
			  
			  StreamResult output=new StreamResult(System.out);
			  transformer.transform(source, output);
        }
	  catch(ParserConfigurationException pe)
	  {
		  System.out.println(pe);  
	  }
	  catch(TransformerException te)
	  {
		  System.out.println(te); 
	  }
	 /* catch(XPathExpressionException xpe)
	  {
		  System.out.println(xpe); 
	  }*/
}



}
