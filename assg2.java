import java.io.BufferedReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.InputStreamReader;
import java.io.Writer;
import java.util.Iterator;


/**
 *
 * @author esp
 */



public class assg2 {
    
    private static final String FILE_IN = "C:\\Users\\esp\\Desktop\\StudentPracticumLists.xlsx";
    private static final String FILE_OUT ="C:\\Users\\esp\\Desktop\\studfile.md";
    
    
    public static void run() {

        Writer w = null;
      
        boolean lbreak = true;
    

        try {

            DataFormatter data = new DataFormatter();

            FileInputStream excellist = new FileInputStream(new File(FILE_IN));
            Workbook workbook = new XSSFWorkbook(excellist);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            File f = new File(FILE_OUT);
            w = new BufferedWriter(new FileWriter(f));

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();
                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    String d = data.formatCellValue(currentCell);
                    w.write(d + " ");
                }
                w.write("\r\n");
                if (lbreak == true) {

                    lbreak = false;
                }

            }

        } catch (FileNotFoundException e) {
        } catch (IOException e) {
        }

        try {
            if (w != null) {
                w.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static void main(String[] args) throws IOException
    {
        File file = new File ("C:\\Users\\esp\\Desktop\\studfile.md");
        FileInputStream fileStream = new FileInputStream(file);
        InputStreamReader input = new InputStreamReader(fileStream);
        BufferedReader reader = new BufferedReader(input);
         
        String line;
         
        int countWord = 0;
        int sentenceCount = 0;
        int characterCount = 0;
        int paragraphCount = 1;
        
         
        while((line = reader.readLine()) != null)
        {
            
            if(line.equals(""))
            {
                paragraphCount++;
            }
            if(!(line.equals("")))
            {
                 
                
                // \\s+ is the space delimiter in java
                String[] wordList = line.split("\\s+");
                
                 characterCount += line.length();
               
                 
                countWord += wordList.length;
                 
                String[] sentenceList = line.split("[!?.:]+");
                 
                sentenceCount += sentenceList.length;
                
                
            }
                
        }
         
        System.out.println("Total word count = " + countWord);
        System.out.println("Total number of sentences = " + sentenceCount);
        System.out.println("Total number of characters = " + characterCount);
          
        
        
        
    }   
}