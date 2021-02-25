package vn.com.nsmv.vinhnq;

import java.io.File;
import java.io.IOException;
import java.net.ConnectException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.io.FilenameUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.artofsolving.jodconverter.DocumentFormat;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;




public class DocumentConverter {
	private Config config;
	
	public Config getConfig() {
		return config;
	}

	public void setConfig(Config config) {
		this.config = config;
	}
	protected static final Logger logger = LoggerFactory.getLogger(DocumentConverter.class);
/**
	 * yyyy/MM/dd HH:mm:ss
	 */
	public static final SimpleDateFormat formatterYYYYMMDDHHMMss = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
	
	/**
	 * yyyyMMddHHmmssSSS
	 */
	public static final SimpleDateFormat formatterYYYYMMDDHHMMSSSSS = new SimpleDateFormat("yyyyMMddHHmmssSSS");
	
	public DocumentConverter(Config config) {
		this.config = config;
	}
	
	public DocumentConverter() {
		super();
	}

	/**
	 * @param inputFileByteArray
	 * @param inputExtension
	 * @param outputExtension
	 * @return byte[]
	 * @throws Exception
	 */
	public byte[] convert(byte[] inputFileByteArray, String inputExtension, String outputExtension) throws Exception {
		String tmpFileName = formatterYYYYMMDDHHMMSSSSS.format(new Date());
		if(null == this.config) {
			throw new NullPointerException("config cannot be null");
		}
		existsAndCreateDirectory(this.config.getOutputTMPDirectory());
		Path inputPath = Paths.get(this.config.getOutputTMPDirectory(), tmpFileName + "." + inputExtension);
		Path outputPath = Paths.get(this.config.getOutputTMPDirectory(), tmpFileName + "." + outputExtension);
		File inputFile = new File(inputPath.toString());
		File outputFile = new File(outputPath.toString());
		org.apache.commons.io.FileUtils.writeByteArrayToFile(inputFile, inputFileByteArray);
		DocumentFormat inputDocumentFormat = BasicDocumentFormat.getFormatByFileExtension(inputExtension);
		DocumentFormat outputDocumentFormat = BasicDocumentFormat.getFormatByFileExtension(outputExtension);
		String outputPathTmp = convert(inputFile, inputDocumentFormat, outputFile, outputDocumentFormat, true).getAbsolutePath();
		byte[] output = Files.readAllBytes(Paths.get(outputPathTmp));
		outputFile.delete();
		return output;
	}
	
	/**
	 * @param inputFile
	 * @param outputExtension
	 * @param deleteInputFile
	 * @return File - Output File
	 */
	public File convert(File inputFile,String outputExtension,boolean deleteInputFile) {
		if(null == this.config) {
			throw new NullPointerException("config cannot be null");
		}
		existsAndCreateDirectory(inputFile.getParent());
		if (!inputFile.exists()) {
			throw new RuntimeException(inputFile.getPath() + System.lineSeparator() + "FILE NOT FOUND");
		}
		String inputExtension = FilenameUtils.getExtension(inputFile.getPath());
		Path outputPath = Paths.get(this.config.getOutputTMPDirectory(), inputFile.getName().replace(inputExtension, outputExtension));
		File outputFile = new File(outputPath.toString());
		DocumentFormat inputDocumentFormat = BasicDocumentFormat.getFormatByFileExtension(inputExtension);
		DocumentFormat outputDocumentFormat = BasicDocumentFormat.getFormatByFileExtension(outputExtension);
		return convert(inputFile, inputDocumentFormat, outputFile, outputDocumentFormat, deleteInputFile);
	}
	
	/**
	 * @param inputFile
	 * @param outputFile
	 * @param deleteInputFile
	 * @return File - Output File
	 */
	public File convert(String inputPath, String outputPath, boolean deleteInputFile) {

		File inputFile = new File(inputPath);
		File outputFile = new File(outputPath);
		if (!inputFile.exists()) {
			throw new RuntimeException(inputFile.getPath() + System.lineSeparator() + "FILE NOT FOUND");
		}
		if (!outputFile.exists()) {
			try {
				outputFile.createNewFile();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		DocumentFormat inputDocumentFormat = BasicDocumentFormat.getFormatByFileExtension(FilenameUtils.getExtension(inputFile.getPath()));
		DocumentFormat outputDocumentFormat = BasicDocumentFormat.getFormatByFileExtension(FilenameUtils.getExtension(outputFile.getPath().toString()));
		return convert(inputFile, inputDocumentFormat, outputFile, outputDocumentFormat, deleteInputFile);
	}
	
	
	/**
	 * @param inputFile
	 * @param outputFile
	 * @param deleteInputFile
	 * @return File - Output File
	 */
	public File convert(File inputFile, File outputFile, boolean deleteInputFile) {
		DocumentFormat inputDocumentFormat = BasicDocumentFormat.getFormatByFileExtension(FilenameUtils.getExtension(inputFile.getPath()));
		DocumentFormat outputDocumentFormat = BasicDocumentFormat.getFormatByFileExtension(FilenameUtils.getExtension(outputFile.getPath().toString()));
		return convert(inputFile, inputDocumentFormat, outputFile, outputDocumentFormat, deleteInputFile);
	}
	/**
	 * @param inputFile
	 * @param outputPath
	 * @param deleteInputFile
	 * @return File - Output File
	 */
	public File convert(File inputFile, Path outputPath, boolean deleteInputFile) {
		existsAndCreateDirectory(outputPath.toString());
		File outputFile = new File(outputPath.toString());
		DocumentFormat inputDocumentFormat = BasicDocumentFormat.getFormatByFileExtension(FilenameUtils.getExtension(inputFile.getAbsolutePath()));
		DocumentFormat outputDocumentFormat = BasicDocumentFormat.getFormatByFileExtension(FilenameUtils.getExtension(outputPath.toString()));
		return convert(inputFile, inputDocumentFormat, outputFile, outputDocumentFormat, deleteInputFile);
	}
	
	/**
	 * @param inputFile
	 * @param inputDocumentFormat - BasicDocumentFormat.{something}
	 * @param outputFile
	 * @param outputDocumentFormat - BasicDocumentFormat.{something}
	 * @param deleteInputFile - Delete input file after convert complete???
	 * @return File - Output File
	 * @implNote - Linux command: libreoffice7.0  --headless --accept="socket,host=127.0.0.1,port=8100;urp;" --nofirststartwizard
	 * @implNote - Windows command: soffice -headless -nologo -nodefault -invisible -nofirststartwizard -norestore -accept=socket,host=127.0.0.1,port=8100;urp
	 */
	public File convert(File inputFile, DocumentFormat inputDocumentFormat, File outputFile, DocumentFormat outputDocumentFormat, boolean deleteInputFile) {
		existsAndCreateDirectory(inputFile.getParent());
		existsAndCreateDirectory(outputFile.getParent());
		OpenOfficeConnection connection = null;
		com.artofsolving.jodconverter.DocumentConverter converter = null;
		if(null == inputDocumentFormat) {
			throw new NullPointerException("inputDocumentFormat Extension");
		}
		if(null == outputDocumentFormat) {
			throw new NullPointerException("outputDocumentFormat Extension");
		}
		try {
			if (!inputFile.exists()) {
				throw new RuntimeException(inputFile.getPath() + System.lineSeparator() + "FILE NOT FOUND");
			}
			if (!inputFile.canRead()) {
				throw new RuntimeException(inputFile.getPath() + System.lineSeparator() + "FILE CAN'T READ");
			}

			if (outputFile.exists()) {
				// A file with the same name already exists, so delete it
				if (!outputFile.delete()) {
					throw new RuntimeException(outputFile.getPath() + System.lineSeparator() + "OUTPUT FILE ALREADY EXISTED - TRY THE DELETION BUT FAILED");
				}
			}
		  
			//
			//START CONVERSION
			connection = new SocketOpenOfficeConnection("127.0.0.1", 8100);
			connection.connect();
			converter = new OpenOfficeDocumentConverter(connection);
			converter.convert(inputFile, inputDocumentFormat, outputFile, outputDocumentFormat);
			connection.disconnect();

			if (!outputFile.exists()) {
				throw new RuntimeException(outputFile.getPath() + System.lineSeparator() + "THE OUTPUT FILE HAS NOT BEEN CREATED EVEN THOUGH THE CONVERSION IS COMPLETE");
			}
			if (!outputFile.canRead()) {
				throw new RuntimeException(outputFile.getPath() + System.lineSeparator() + "THE OUTPUT FILE CAN'T READ");
			}
			System.err.println("===========================================================");
			System.err.println("INPUT: " + inputFile.getPath());
			System.err.println("OUTPUT: " + outputFile.getPath());
			System.err.println("===========================================================");
			//
			// DELETE INPUT FILE
			if (deleteInputFile) {
				inputFile.delete();
			}
			return outputFile;
		} catch (ConnectException ex) {
			String mess = "Cannot connect to print service" + System.lineSeparator();
			mess += "Linux command: libreoffice7.0  --headless --accept=\"socket,host=127.0.0.1,port=8100;urp;\" --nofirststartwizard" + System.lineSeparator();
			mess += "Windows command: soffice -headless -nologo -nodefault -invisible -nofirststartwizard -norestore -accept=socket,host=127.0.0.1,port=8100;urp" + System.lineSeparator();
		//	logger.error(ex.toString());
			throw new RuntimeException(mess + ex.toString());
		} catch (Exception ex) {
			// 標準の例外エラー
			System.err.println(ex);
		//	logger.error(ex.toString());
			throw new RuntimeException(ex);
		}
	}
	
	private static boolean existsAndCreateDirectory(String pathStr) {
		Path path = Paths.get(pathStr);
		boolean exists = Files.exists(path);
		try {
	        if(!exists) {
	            Files.createDirectories(path);
	        }
	        return exists;
		} catch (Exception e) {
			System.err.println(e);
			throw new RuntimeException(e);
		}
	}

}
