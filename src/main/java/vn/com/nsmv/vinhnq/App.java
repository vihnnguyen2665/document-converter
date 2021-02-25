package vn.com.nsmv.vinhnq;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Hello world!
 *
 */
public class App 
{
	protected static final Logger logger = LoggerFactory.getLogger(App.class);
	/**
	 * @param input path File
	 * @param output path File
	 * @return File - Output File
	 */
    public static void main( String[] args )
    {
    	if(args.length != 2) {
    		String msg = "Incorrect syntax: " + System.lineSeparator();
    		msg += "document-converter.jar 'input file Path' 'output file path' " + System.lineSeparator();
    		msg += "eg: document-converter.jar 'C:\\sample.xlsx' 'C:\\sample.pdf' ";
    		logger.error(msg);
    		return;
    	}
    	Config config = new Config();
    	config.setOutputTMPDirectory("./");
    	DocumentConverter converter = new DocumentConverter(config);
    	
    	converter.convert( args[0], args[1], false);
    }
}
