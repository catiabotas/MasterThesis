package WekaClassification;

import java.io.File;
import java.io.FilenameFilter;

import weka.core.Instances;
import weka.core.converters.ArffSaver;
import weka.core.converters.CSVLoader;

public class ConvertCSVtoARFF {

	public static void main(String[] args) throws Exception{
		
		// Choose here the folder name to convert from [*.csv] to [*.arff]
		// Note that files need to be [*.csv]
		String folderName = "CSV Files";
		
		// Go through all [*.csv] files in "folderName" folder
		File dir = new File("C:\\Users\\BotasC1\\Desktop\\Tese\\Data\\ProcessedData\\ClassFiles_Update\\ReumaBImp\\" + folderName);
		File[] directoryListing = dir.listFiles(new FilenameFilter() {
	        public boolean accept(File dir, String name) {
	            return name.toLowerCase().endsWith(".csv");
	        }
		});
		if (directoryListing != null) {
			  for (File child : directoryListing) {
				  // ======== File Conversion =============
				  // Load CVS
				  CSVLoader loader = new CSVLoader();
				  loader.setSource(child);
				  Instances data = loader.getDataSet(); // Get instances object
				
				  // Save ARFF
				  ArffSaver saver = new ArffSaver();
				  saver.setInstances(data); // Set the data we want to convert
				  saver.setFile(new File("C:\\Users\\BotasC1\\Desktop\\Tese\\Data\\ProcessedData\\ClassFiles_Update\\ReumaBImp\\" + folderName + "_arff/" + child.getName().replace(".csv",".arff")));
				  //saver.setDestination(new File("C:/Users/Utilizador/Google Drive/IST/Tese/DadosReuma/Trabalho nos dados/Weka/WekaJavaClassification/Data/" + folderName + "_arff/" + child.getName().replace(".csv",".arff")));		
				  saver.writeBatch();
			  }
		}
	}

}
