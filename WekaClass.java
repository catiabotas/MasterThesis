package WekaClassification;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.Arrays;
import java.util.Random;

import weka.classifiers.Evaluation;
import weka.classifiers.bayes.NaiveBayes;
import weka.classifiers.bayes.BayesNet;
import weka.classifiers.functions.Logistic;
import weka.classifiers.functions.SMO;
import weka.classifiers.trees.J48;
import weka.classifiers.trees.RandomForest;
import weka.attributeSelection.*;
import weka.core.Instances;
import weka.filters.Filter;
import weka.filters.supervised.attribute.Discretize;
import weka.filters.supervised.instance.SMOTE;

public class WekaClass {
	
	static String initClassBalStr = ""; // String with information of initial class balance
	static String finalClassBalStr = ""; // String with information of class balance after SMOTE
	
	// Loads a given ARFF file and sets the class attribute as the last one
	// Saves the class balance information in variable [initClassBaltStr]
	protected static Instances loadFile(File child) throws Exception {
		
	     Instances result;
	     BufferedReader reader;
	 
	     reader = new BufferedReader(new FileReader(child));
	     result = new Instances(reader);
	     result.setClassIndex(result.numAttributes() - 1); // class label is last attribute
	     int[] countLabels = result.attributeStats(result.numAttributes() - 1).nominalCounts; // counts total number of data of each class (C0,C1,C2)
	     int totalLabels = result.attributeStats(result.numAttributes() - 1).totalCount; // counts total number of attributes in class label
	     initClassBalStr = initClassBalStr + "\n" + child.getName().replace(".arff","").replace("classFile_", "") + "\t" + String.valueOf(countLabels[1]) + "\t" + String.valueOf(countLabels[0]) + "\t" + String.valueOf(countLabels[2]) + "\t" + String.valueOf(totalLabels) + "\t" + String.format("%.2f",(Float.valueOf(countLabels[1])/Float.valueOf(totalLabels))*100) + " %" + "\t" + String.format("%.2f",(Float.valueOf(countLabels[0])/Float.valueOf(totalLabels))*100) + " %" + "\t" + String.format("%.2f",(Float.valueOf(countLabels[2])/Float.valueOf(totalLabels))*100) + " %";
	     reader.close();
	 
	     return result;
	}
	
	// Complete preprocessing of the data
	protected static Instances preProcessing(Instances result, File child) throws Exception{
		
		result.setClassIndex(result.numAttributes() - 1); // class label is last attribute
	    int[] countLabels = result.attributeStats(result.numAttributes() - 1).nominalCounts; // counts total number of data of each class (C0,C1,C2)
	    int totalLabels = result.attributeStats(result.numAttributes() - 1).totalCount; // counts total number of attributes in class label	     
		
	    // Ordering the indexes of [countLabels] from maximum to minimum number of class data.
	    int[] auxArray = new int[3];
	    Arrays.fill(auxArray, -1);
	    if ((countLabels[0] > countLabels[1]) && (countLabels[0] > countLabels[2])){
	    	auxArray[0] = 0;
	    	if (countLabels[1] > countLabels[2]){
	    		auxArray[1] = 1;
	    		auxArray[2] = 2;
	    	} else {
	    		auxArray[1] = 2;
	    		auxArray[2] = 1;
	    	}
	    } else if ((countLabels[1] > countLabels[0]) && (countLabels[1] > countLabels[2])){
	    	auxArray[0] = 1;
	    	if (countLabels[0] > countLabels[2]){
	    		auxArray[1] = 0;
	    		auxArray[2] = 2;
	    	} else {
	    		auxArray[1] = 2;
	    		auxArray[2] = 0;
	    	}
	    } else if ((countLabels[2] > countLabels[0]) && (countLabels[2] > countLabels[1])){
	    	auxArray[0] = 2;
	    	if (countLabels[0] > countLabels[1]){
	    		auxArray[1] = 0;
	    		auxArray[2] = 1;
	    	} else {
	    		auxArray[1] = 1;
	    		auxArray[2] = 0;
	    	}
	    }
	    
	    // Balance the minimum class
		// SMOTE (Synthetic Minority Oversampling Technique)
		SMOTE smoteFilter1 = new SMOTE();
		smoteFilter1.setInputFormat(result);
		smoteFilter1.setPercentage((Float.valueOf(countLabels[auxArray[0]])/Float.valueOf(countLabels[auxArray[2]])-1.0)*100.0);
        Instances smoteTrain1 = Filter.useFilter(result, smoteFilter1); // apply filter
        
        // Use [smoteTrain1] as train data
        result = smoteTrain1;
        
        // Balance the middle class
		// SMOTE (Synthetic Minority Oversampling Technique)
		SMOTE smoteFilter2 = new SMOTE();
		smoteFilter2.setInputFormat(result);
		smoteFilter2.setPercentage((Float.valueOf(countLabels[auxArray[0]])/Float.valueOf(countLabels[auxArray[1]])-1.0)*100.0);
		smoteFilter2.setClassValue(String.valueOf(auxArray[1]+1));
	    Instances smoteTrain2 = Filter.useFilter(result, smoteFilter2); // apply filter
	     
	    // Use [smoteTrain2] as train data
	    result = smoteTrain2;
		
		// DISCRETIZATION with Fayyad & Irani's MDL method (the default)
		Discretize discreteFilter = new Discretize(); // setup filter
		discreteFilter.setInputFormat(result);
	    Instances discreteTrain = Filter.useFilter(result, discreteFilter); // apply filter
	    
	    // Use [discreteTrain] as train data
	    result = discreteTrain;
	    
	    // Save information of the preprocessed data
	    countLabels = result.attributeStats(result.numAttributes() - 1).nominalCounts; // counts total number of data of each class (C0,C1,C2)
	    totalLabels = result.attributeStats(result.numAttributes() - 1).totalCount; // counts total number of attributes in class label
	    finalClassBalStr = finalClassBalStr + "\n" + child.getName().replace(".arff","").replace("classFile_", "") + "\t" + String.valueOf(countLabels[1]) + "\t" + String.valueOf(countLabels[0]) + "\t" + String.valueOf(countLabels[2]) + "\t" + String.valueOf(totalLabels) + "\t" + String.format("%.2f",(Float.valueOf(countLabels[1])/Float.valueOf(totalLabels))*100) + " %" + "\t" + String.format("%.2f",(Float.valueOf(countLabels[0])/Float.valueOf(totalLabels))*100) + " %" + "\t" + String.format("%.2f",(Float.valueOf(countLabels[2])/Float.valueOf(totalLabels))*100) + " %";
		
		return result;
	}
	
	// Saves a specific [data] in a specific [fileName].txt file in folder "files"
	protected static void saveDateTextFile(String fileName, String data) throws IOException{
	    FileWriter newFile = new FileWriter("files/" + fileName + ".txt"); // Create file 
	    BufferedWriter out = new BufferedWriter(newFile);
	    out.write(data); // write in the file
	    out.close(); // close the file
	}
	

	  //Uses Filter
	  protected static Instances useFilter(Instances data) throws Exception {
	    weka.filters.supervised.attribute.AttributeSelection filter = new weka.filters.supervised.attribute.AttributeSelection();
	    CfsSubsetEval eval = new CfsSubsetEval();
	    BestFirst search = new BestFirst();
	    filter.setEvaluator(eval);
	    filter.setSearch(search);
	    filter.setInputFormat(data);
	    Instances newData = Filter.useFilter(data, filter);
	    return newData;
	  }
	
	
	// Main method
	public static void main(String[] args) throws Exception{
		
		// ==========================================================
		// CHOOSE HERE THE FOLDER NAME WITH FILES TO CLASSIFY
		// Note that files need to be ARFF
		String folderName = "CSVFiles_arff";
		// ==========================================================
		
		// path to folder
		File dir = new File("C:\\Users\\BotasC1\\Desktop\\Tese\\Data\\ProcessedData\\ClassFiles_Update\\ReumaBImp\\" + folderName);
		// all ARFF files from directory are considered
		File[] directoryListing = dir.listFiles(new FilenameFilter() {
	        public boolean accept(File dir, String name) {
	            return name.toLowerCase().endsWith(".arff");
	        }
		});
		// check if there are files in the directory
		if (directoryListing != null) {
			// names of the columns of the performance table
			String outStr = "data\t\t\tNaiveBayes\tBayesNet\tLogistic\tSMO\t\tJ48\t\tRandomForest";
			String outAtr = "";
			String outStrFS = "data\t\t\tNaiveBayes\tBayesNet\tLogistic\tSMO\t\tJ48\t\tRandomForest";
			
			// names of the columns of class balance table
			initClassBalStr = "data\t\t\tC0\tC1\tC2\tTotal\tC0(%)\tC1(%)\tC2(%)";
			finalClassBalStr = "data\t\t\tC0\tC1\tC2\tTotal\tC0(%)\tC1(%)\tC2(%)";
			
			// counter of the files in the directory
			int count = 1;
			
			for (File child : directoryListing) {
				
				// ======== GET DATA FROM ARFF FILE =========
				// reading file and creating instances for classification where class label is the last attribute
				Instances train = loadFile(child);
				
				// ======== PRE-PROCESSING OF DATA =========
				Instances trainProcessed = preProcessing(train, child);
				
			    // Use [trainProcessed] as train data
				train = trainProcessed;
			    
				// ======== WEKA Classification =============
				
				// Choosing the class label (last column - [cod_resposta_label])
				train.setClassIndex(train.numAttributes()-1);
				
				// Naive Bayes classifier
				NaiveBayes nB = new NaiveBayes();
				nB.buildClassifier(train);
				Evaluation evalNB = new Evaluation(train);
				evalNB.crossValidateModel(nB, train, 10, new Random(1));
				System.out.println("NaiveBayes " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsNaiveBayes_"), evalNB.toSummaryString("\nResultsNaiveBayes\n======\n",true)+"\n"+evalNB.toClassDetailsString()+"\n"+evalNB.toMatrixString());
				
				// Bayes Net classifier
				BayesNet bN = new BayesNet();
				bN.buildClassifier(train); // discrete data
				Evaluation evalBN = new Evaluation(train);
				evalBN.crossValidateModel(bN, train, 10, new Random(1));
				System.out.println("BayesNet " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsBayesNet_"), evalBN.toSummaryString("\nResultsBayesNet\n======\n",true)+"\n"+evalBN.toClassDetailsString()+"\n"+evalBN.toMatrixString());
				
				// Logistic classifier
				Logistic logi = new Logistic();
				logi.buildClassifier(train);
				Evaluation evalLogi = new Evaluation(train);
				evalLogi.crossValidateModel(logi, train, 10, new Random(1));
				System.out.println("Logistic " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsLogistic_"), evalLogi.toSummaryString("\nResultsLogistic\n======\n",true)+"\n"+evalLogi.toClassDetailsString()+"\n"+evalLogi.toMatrixString());
				
				// SMO (PolyKernel) classifier
				SMO smo = new SMO();
				smo.buildClassifier(train);
				Evaluation evalSMO = new Evaluation(train);
				evalSMO.crossValidateModel(smo, train, 10, new Random(1));
				System.out.println("SMO " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsSMO_"), evalSMO.toSummaryString("\nResultsSMO\n======\n",true)+"\n"+evalSMO.toClassDetailsString()+"\n"+evalSMO.toMatrixString());
				
				// J48 classifier
				J48 j48 = new J48();
				j48.buildClassifier(train);
				Evaluation evalJ48 = new Evaluation(train);
				evalJ48.crossValidateModel(j48, train, 10, new Random(1));
				System.out.println("J48 " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsJ48_"), evalJ48.toSummaryString("\nResultsJ48\n======\n",true)+"\n"+evalJ48.toClassDetailsString()+"\n"+evalJ48.toMatrixString()+"\n"+j48.graph());
				
				//RandomForest Classifier
				RandomForest rf = new RandomForest();
		        rf.buildClassifier(train);
		        Evaluation evalrf = new Evaluation(train);
		        evalrf.crossValidateModel(rf, train, 10, new Random(1));
		        System.out.println("RandomForest " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsRandomForest_"), evalrf.toSummaryString("\nResultsRandomForest\n======\n",true)+"\n"+evalrf.toClassDetailsString()+"\n"+evalrf.toMatrixString());
				
				// Output string
				outStr = outStr + "\n" + child.getName().replace(".arff","").replace("classFile_", "") + "\t" + String.format("%.3f", evalNB.pctCorrect()) + " %\t" + String.format("%.3f", evalBN.pctCorrect()) + " %\t" + String.format("%.3f", evalLogi.pctCorrect()) + " %\t" + String.format("%.3f", evalSMO.pctCorrect()) + " %\t" + String.format("%.3f", evalJ48.pctCorrect()) + " %\t" + String.format("%.3f", evalrf.pctCorrect()) + " %";
				
				// ======== Select Attributes =============
				
				//Select Attributes
				AttributeSelection atrib = new AttributeSelection();
			    CfsSubsetEval eval = new CfsSubsetEval();
			    BestFirst search = new BestFirst();
			    atrib.setEvaluator(eval);
			    atrib.setSearch(search);
				atrib.SelectAttributes(train);
				System.out.println("Select Attributes " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultSelectAttributes_"), atrib.toResultsString());
				
				//OutpuAttributes
				outAtr = outAtr + "\n" + child.getName().replace(".arff","").replace("classFile_", "") + "\n" + atrib.toResultsString();
				
				// ======== WEKA Classification FeatureSelection =============
				
				// train with FeatureSelection
				Instances trainFeatureSelection = useFilter(train);
				
				// train is new data with Feature Selection
				train = trainFeatureSelection;
				
				// Naive Bayes classifier
				NaiveBayes nB_FS = new NaiveBayes();
				nB_FS.buildClassifier(train);
				Evaluation evalNB_FS = new Evaluation(train);
				evalNB_FS.crossValidateModel(nB_FS, train, 10, new Random(1));
				System.out.println("NaiveBayes FeatureSelection " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsNaiveBayes_FeatureSelection_"), evalNB_FS.toSummaryString("\nResultsNaiveBayes_FeatureSelection\n======\n",true)+"\n"+evalNB_FS.toClassDetailsString()+"\n"+evalNB_FS.toMatrixString());
				
				// Bayes Net classifier
				BayesNet bN_FS = new BayesNet();
				bN_FS.buildClassifier(train); // discrete data
				Evaluation evalBN_FS = new Evaluation(train);
				evalBN_FS.crossValidateModel(bN_FS, train, 10, new Random(1));
				System.out.println("BayesNet FeatureSelection " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsBayesNet_FeatureSelection_"), evalBN_FS.toSummaryString("\nResultsBayesNet_FeatureSelection\n======\n",true)+"\n"+evalBN_FS.toClassDetailsString()+"\n"+evalBN_FS.toMatrixString());
				
				// Logistic classifier
				Logistic logiFS = new Logistic();
				logiFS.buildClassifier(train);
				Evaluation evalLogiFS = new Evaluation(train);
				evalLogiFS.crossValidateModel(logiFS, train, 10, new Random(1));
				System.out.println("Logistic FeatureSelection " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsLogistic_FeatureSelection_"), evalLogiFS.toSummaryString("\nResultsLogistic_FeatureSelection\n======\n",true)+"\n"+evalLogiFS.toClassDetailsString()+"\n"+evalLogiFS.toMatrixString());
				
				// SMO (PolyKernel) classifier
				SMO smoFS = new SMO();
				smoFS.buildClassifier(train);
				Evaluation evalSMO_FS = new Evaluation(train);
				evalSMO_FS.crossValidateModel(smoFS, train, 10, new Random(1));
				System.out.println("SMO FeatureSelection " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsSMO_FeatureSelection_"), evalSMO_FS.toSummaryString("\nResultsSMO_FeatureSelection\n======\n",true)+"\n"+evalSMO_FS.toClassDetailsString()+"\n"+evalSMO_FS.toMatrixString());
				
				// J48 classifier
				J48 j48_FS = new J48();
				j48_FS.buildClassifier(train);
				Evaluation evalJ48_FS = new Evaluation(train);
				evalJ48_FS.crossValidateModel(j48_FS, train, 10, new Random(1));
				System.out.println("J48 FeatureSelection " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsJ48_FeatureSelection_"), evalJ48_FS.toSummaryString("\nResultsJ48_FeatureSelection\n======\n",true)+"\n"+evalJ48_FS.toClassDetailsString()+"\n"+evalJ48_FS.toMatrixString()+"\n"+j48_FS.graph());
				
				//RandomForest Classifier
				RandomForest rf_FS = new RandomForest();
		        rf_FS.buildClassifier(train);
		        Evaluation evalrf_FS = new Evaluation(train);
		        evalrf_FS.crossValidateModel(rf_FS, train, 10, new Random(1));
		        System.out.println("RandomForest FeatureSelection " + count + " done");
				// saves information in file
				saveDateTextFile(child.getName().replace(".arff","").replace("classFile_", "resultsRandomForest_FeatureSelection_"), evalrf_FS.toSummaryString("\nResultsRandomForest_FeatureSelection\n======\n",true)+"\n"+evalrf_FS.toClassDetailsString()+"\n"+evalrf_FS.toMatrixString());
				
				// Output string
				outStrFS = outStrFS + "\n" + child.getName().replace(".arff","").replace("classFile_", "") + "\t" + String.format("%.3f", evalNB_FS.pctCorrect()) + " %\t" + String.format("%.3f", evalBN_FS.pctCorrect()) + " %\t" + String.format("%.3f", evalLogiFS.pctCorrect()) + " %\t" + String.format("%.3f", evalSMO_FS.pctCorrect()) + " %\t" + String.format("%.3f", evalJ48_FS.pctCorrect()) + " %\t" + String.format("%.3f", evalrf_FS.pctCorrect()) + " %";
				count++;
				
			}
			
			// Print table with results
			System.out.println("");
			System.out.println("========== Results without FeatureSelection ==========");
			System.out.println(outStr);
			
			// Print table with results
			System.out.println("");
			System.out.println("========== Results Select Attributes ==========");
			System.out.println(outAtr);
			
			// Print table with results
			System.out.println("");
			System.out.println("========== Results Feature Selection ==========");
			System.out.println(outStrFS);
			
			// saves information regarding performances in file
			saveDateTextFile("resultsAllPerformances", outStr);
			saveDateTextFile("resultsAllSelectedAttributes", outAtr);
			saveDateTextFile("resultsAllPerformances_FeatureSelection", outStrFS);
			
			// saves information regarding class balance in file (before and after preprocessing)
			saveDateTextFile("ClassBalanceBefore", initClassBalStr);
			saveDateTextFile("ClassBalanceAfter", finalClassBalStr);
			//System.out.println("========== ClassBalance ==========");
			//System.out.println(countStr);
		}
	}
	
}
