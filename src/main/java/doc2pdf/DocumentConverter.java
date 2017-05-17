package doc2pdf;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.jodconverter.OfficeDocumentConverter;
import org.jodconverter.document.DocumentFormat;
import org.jodconverter.document.DocumentFormatRegistry;
import org.jodconverter.filter.DefaultFilterChain;
import org.jodconverter.filter.RefreshFilter;
import org.jodconverter.office.DefaultOfficeManagerBuilder;
import org.jodconverter.office.OfficeManager;

public class DocumentConverter {
	protected static OfficeManager officeManager;
	protected static OfficeDocumentConverter converter;
	protected static DocumentFormatRegistry formatRegistry;
	  
	public static void main(String[] args) throws Exception {
		File inputFile = new File("test.docx");

		officeManager = new DefaultOfficeManagerBuilder().build();
    converter = new OfficeDocumentConverter(officeManager);
    formatRegistry = converter.getFormatRegistry();

    officeManager.start();
    
    final DefaultFilterChain chain = new DefaultFilterChain(RefreshFilter.INSTANCE);
    
    convertFilePDF(inputFile, null, chain);
    
    officeManager.stop();
	}
	
	protected static void convertFilePDF(final File inputFile, final File outputDir, final DefaultFilterChain chain) throws Exception {
    DocumentFormat outputFormat = formatRegistry.getFormatByExtension("pdf");
    File outputFile = null;
    
    if (outputDir == null) {
      outputFile = new File(FilenameUtils.getBaseName(inputFile.getName()) + "." + outputFormat.getExtension());
    } else {
      outputFile = new File(outputDir, FilenameUtils.getBaseName(inputFile.getName()) + "." + outputFormat.getExtension());
      FileUtils.deleteQuietly(outputFile);
    }
    
    converter.convert(chain, inputFile, outputFile, outputFormat);
    
    chain.reset();
  }
}
