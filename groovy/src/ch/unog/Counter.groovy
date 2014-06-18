package ch.unog


import org.apache.tika.sax.BodyContentHandler
import org.xml.sax.helpers.DefaultHandler
import org.apache.tika.config.TikaConfig
import org.apache.tika.metadata.Metadata
import org.apache.tika.parser.AutoDetectParser
import org.apache.tika.parser.microsoft.OfficeParser
import org.apache.tika.parser.ParseContext
import org.apache.tika.parser.Parser
import org.apache.tika.metadata.TikaCoreProperties;
 
import static org.apache.tika.metadata.TikaCoreProperties.*;
  
import groovy.util.logging.Slf4j
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.security.MessageDigest


@Slf4j
class Counter {
	/*
	 * Given a filename (full path to file), returns the MS Word word count
	 */
	
	public int count(String fileName) {
		def mde = new MetadataExtractor()
		def props = new Properties()
		new File("wordcounter.properties").withInputStream {
			stream -> props.load(stream)
		 }
		int maxfilesize = props['maxfilesizeMB'].toInteger() *1024*1024

		def file = new File(fileName)
		int cnt

		if (file.exists()){
			if (file.length() < maxfilesize){
				log.info("about to parse: ${file.name}")
				Map metaDataFields = mde.getMetadataForFile(file)
				//log.info(metaDataFields.toMapString())
	 			cnt = metaDataFields['Word-Count']?.toInteger()?:0
				log.info(cnt.toString())
			} else {
				log.warn("file too large - skipped: ${file.name}")
			}
		}
		mde = null
		file = null
		return cnt
	}
		
	static main(args) {
		def obj = new Counter()
		log.info('starting counter')
		def result = obj.count(args[0])
		println result
	}

	
	
	/**
	 * Taken from
	 * https://gist.github.com/kaisternad/7736686
	 *
	 */
	@Slf4j
    static class MetadataExtractor{

        /**
         * List of Metadata fields to be extracted.
         * Change these if you would like to extract different fields
         */
        private static final List METADATA_FIELDS = [
            Metadata.CONTENT_TYPE,
			Metadata.WORD_COUNT,
			Metadata.PAGE_COUNT,
			Metadata.AUTHOR
        ]

        public Map getMetadataForFile(File file){
            Metadata metadata = parseFile(file)
            String md5 = generateMD5(file)
            Map extractedMetadata = extractMetadata(metadata)
            extractedMetadata << ["file-md5" : md5]
            return extractedMetadata
        }

        private Map extractMetadata(Metadata tikaMeta){
            def nonEmptyFields = [:]

            METADATA_FIELDS.each{ field ->
                def extractedMetadataField = tikaMeta.get(field);
                if (extractedMetadataField){
                    String key = (field.class.equals(String.class) ? field : field.name).toString()
                    nonEmptyFields << [(key):extractedMetadataField]
                    
                }
            }
            return nonEmptyFields;
        }
        
        private String generateMD5(File f) {
            MessageDigest digest = MessageDigest.getInstance("MD5")
            digest.update(f.getBytes());
            new BigInteger(1, digest.digest()).toString(16).padLeft(32, '0')
        }
        
        private Metadata parseFile(File file){
            FileInputStream stream = new FileInputStream(file);
            TikaConfig tikaConfig = new TikaConfig()
            Metadata tikaMeta = new Metadata()
			//BodyContentHandler handler = new BodyContentHandler(10*1024*1024);
			DefaultHandler handler = new DefaultHandler();
            Parser parser = new AutoDetectParser(tikaConfig)
			//Parser parser = new OfficeParser()
			ParseContext pc = new ParseContext()
            try {
                parser.parse(stream, handler, tikaMeta, pc)
                log.debug("parsed file {$file.absolutePath}")
            } catch (Exception e) {
                log.error("Failed to parse file ${file.absolutePath}  ${e}")
            }
			stream = null
			handler = null
            return tikaMeta
        }
    }
		
}
