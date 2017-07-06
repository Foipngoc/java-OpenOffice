package com.open;

import java.io.File;
import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

public class office 
{
	public   static  void  main(String [] args)
	{
		
		office2PDF("D:/11.docx","D:/dd.pdf");		
	}
	
	public static int office2PDF(String sourceFile, String destFile)
	{  
        try {  
            File inputFile = new File(sourceFile);  
            if (!inputFile.exists()) {  
                return -1; 
            }  
  
            File outputFile = new File(destFile);  
            if (!outputFile.getParentFile().exists()) {  
                outputFile.getParentFile().mkdirs();  
            }  
               
            OpenOfficeConnection connection = new SocketOpenOfficeConnection(  
                    "127.0.0.1", 8100);  
            connection.connect();  
  
            // convert  
            DocumentConverter converter = new OpenOfficeDocumentConverter(  
                    connection);  
            converter.convert(inputFile, outputFile);  
  
            // close the connection  
            connection.disconnect();  
  
            return 0;  
        } catch (Exception e) 
        {  
            e.printStackTrace();  
            return -1;  
        }  
    } 

}
