import org.junit.Test;
import static org.junit.Assert.*;

import java.io.IOException;

import org.unisens.UnisensParseException;
import org.unisens.unisens2excel.*;

/*
 * This Java source file was auto generated by running 'gradle init --type java-library'
 * by 'juergen.stumpp' at '21.07.16 18:03' with Gradle 2.14.1
 *
 * @author juergen.stumpp, @date 21.07.16 18:03
 */
public class LibraryTest {
    
    @Test public void testXlsRendering() throws UnisensParseException, IOException {
    	
   		Unisens2Excel u2xls= new Unisens2Excel("./src/test/resources/UnisensTestData/T1", 1.0/60, "./Results1.xlsx");
  		u2xls.renderXLS();
    }
    
    @Test public void testXlsRendering2() throws UnisensParseException, IOException {
   
   		Unisens2Excel u2xls= new Unisens2Excel("./src/test/resources/UnisensTestData/T2", 1.0/10, "./Results2.xlsx");
  		u2xls.renderXLS();
    }
    
    @Test public void testXlsRendering3A() throws UnisensParseException, IOException {
    	   
   		Unisens2Excel u2xls= new Unisens2Excel("./src/test/resources/UnisensTestData/T3", 1.0/10, "./Results3A.xlsx");
  		u2xls.renderXLS();
    }
    
    @Test public void testXlsRendering3B() throws UnisensParseException, IOException {
 	   
   		Unisens2Excel u2xls= new Unisens2Excel("./src/test/resources/UnisensTestData/T3", 1.0/10, "./Results3B.xlsx");
   		u2xls.setMarkerFormat(MarkerFormat.SIMPLE);
  		u2xls.renderXLS();
    }
    
    
    @Test public void testBatch() throws UnisensParseException, IOException {
    	Unisens2Excel.batchProcess("./src/test/resources/UnisensTestData", 1.0/60);
    }
   
}
