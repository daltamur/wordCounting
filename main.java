import org.apache.commons.compress.utils.FileNameUtils;
import org.apache.commons.io.comparator.NameFileComparator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFComment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.*;
import java.util.regex.Pattern;

import static org.apache.commons.io.comparator.ExtensionFileComparator.EXTENSION_SYSTEM_COMPARATOR;

public class main {
    public  static void main(String[] args) throws IOException, InvalidFormatException {
        //Gets a directory from the user & makes sure it is a valid directory
        Scanner scanner=new Scanner(System.in);
        System.out.println("Enter a source directory:");
        String source=scanner.nextLine();
        File directory=new File(source);
        if(!directory.exists()){
            System.out.println("No such directory");
            main(args);
        }
        //Get a list of all the .docx extension files with the brackets used to indicate a name tag
        ArrayList<File>unnestedFiles=new ArrayList<>();
        getAlldocXFiles(directory, unnestedFiles);
        HashMap<String,HashMap<String,studentCounts>> student=new HashMap<>();
        File[]filesArray=unnestedFiles.toArray(new File[0]);
        //make the csv file
        wordCounter(filesArray,source,student);
    }

    private static void getAlldocXFiles(File directory, ArrayList<File> fileList) throws InvalidFormatException, IOException {
        //if the current file is not a directory and has a docx extension, add it to the list of files to be processed
        for(File currentFile: Objects.requireNonNull(directory.listFiles())){
            if(currentFile.isFile()&& FileNameUtils.getExtension(currentFile.getName()).equals("docx")){
                XWPFDocument currentDoc = new XWPFDocument(OPCPackage.open(currentFile.getAbsolutePath()));
                XWPFWordExtractor ex = new XWPFWordExtractor(currentDoc);
                if(ex.getText().contains("[")) {
                    fileList.add(currentFile);
                }
                currentDoc.close();
                ex.close();
                //if it IS a directory, check the files inside it
            }else if(currentFile.isDirectory()){
                getAlldocXFiles(currentFile,fileList);
            }
        }

    }

    private static void wordCounter(File[] allFiles,String source,HashMap<String ,HashMap<String,studentCounts>>student) throws InvalidFormatException, IOException {
        //checks every file in the list we made in the main method
        for(File current:allFiles) {
            System.out.println("Analyzing "+current.getName());
            //redundant docx extension checker but just for piece of mind
            if (current.isFile() && FileNameUtils.getExtension(current.getName()).equals("docx")) {
                XWPFDocument currentDoc = new XWPFDocument(OPCPackage.open(current.getAbsolutePath()));
                //get the list of comments made
                XWPFComment[] comments = currentDoc.getComments();
                XWPFWordExtractor ex = new XWPFWordExtractor(currentDoc);
                String doc = ex.getText().trim();
                //take out any comments so ad to not inflate the word count
                for (XWPFComment comment : comments) {
                    doc = doc.replace("Comment by Ulises Mejias: " + comment.getText(), " ");
                }
                //take out any extra white space, only one space between each word
                doc = doc.replaceAll("\\s+", " ");
                //name tag found
                if(doc.contains("[")){
                    //reduce the string of the file to where the brackets start, less analyzing this way
                    int bracket = doc.indexOf('[');
                    doc = doc.substring(bracket);
                    //create an array of the different comment entries in the document
                    String[] currentFileString = doc.split("\\[");
                    //analyze each comment
                    for (int i = 1; i < currentFileString.length; i++) {
                        //get the name of the commenter and make sure it is not the "name tag" initial entry
                        String name = currentFileString[i].substring(0, currentFileString[i].indexOf(']'));
                        if (!name.equals("name tag")) {
                            HashMap<String, studentCounts> docWordCount;
                           //if the student had already been added to the hashmap, just get the hash map holding their word count values
                            if (student.containsKey(name)) {
                                docWordCount = student.get(name);
                            } else {
                                //first time student has been referenced, add them to the hashmap with an empty docWordCount hashmap value
                                docWordCount = new HashMap<>();
                                student.put(name, docWordCount);
                            }
                            //if this is the first time the document has been referenced for the current student, initialize it
                            if(!(docWordCount.containsKey(current.getName()))){
                                docWordCount.put(current.getName(), new studentCounts(0,0));
                            }
                            //get the word count for the student, again removing any extra white space that may exist
                            String currentWordSubstring = currentFileString[i].substring(currentFileString[i].indexOf(']') + 1).replaceAll("\\s+", " ");
                            //if it starts with a space, remove it
                            if (currentWordSubstring.startsWith(" ")) {
                                currentWordSubstring = currentWordSubstring.substring(1);
                            }
                            //get the word count for the student
                            int thisCount = currentWordSubstring.split("\\s+").length;
                            //System.out.println(thisCount);
                            //if the end of the word count includes any dangling words that are not part of their contribution, remove them from the word count number
                            if(Pattern.compile("\\d $").matcher(currentWordSubstring).find()||Pattern.compile("\\d$").matcher(currentWordSubstring).find()||Pattern.compile(" Discussion $").matcher(currentWordSubstring).find()){
                                thisCount--;
                            }
                            //create a studentCounts class reference from the current student
                            studentCounts replace=docWordCount.get(current.getName());
                            //increase word and comment count, put the new value in the hashmap
                            replace.increaseWordCount(thisCount);
                            replace.increaseContributionCount();
                            docWordCount.replace(current.getName(), replace);
                            student.replace(name, docWordCount);

                            //System.out.println(thisCount);
                            //System.out.println(currentWordSubstring);
                            //System.out.println(currentFileString[i].substring(currentFileString[i].indexOf(']')+1).replaceAll("\\s+"," "));
                        }
                        //if (fileLength / currentCount == 2) {
                        //    System.out.println("50% Done Gathering Instances");
                        //} else if (fileLength / currentCount == 1) {
                        //    System.out.println("Finished Gathering Instances");
                        //}

                    }
                    currentDoc.close();
                }else{
                    currentDoc.close();
                }
            }
        }
        //write to CSV file in a new folder holding the Analysis csv
        System.out.println("Writing Analysis to CSV File...");
        File analysisFile=new File(source+"Analysis");
        FileWriter writer = new FileWriter(source+"Analysis/analysis.csv");
        BufferedWriter csvWriter=new BufferedWriter(writer);
        StringBuilder heading= new StringBuilder();
        int finalCount=0;
        for(String key:student.keySet()){
            //get the bare file names to write to the csv
            String[] allFileNames=new String[allFiles.length];
            for(int i=0;i<allFileNames.length;i++){
                allFileNames[i]=allFiles[i].getName();
            }
            //Sort file names according to their ASCII value
            Arrays.sort(allFileNames);
            for(String doc:allFileNames){
                //take out any commas in the file name that we'll write so the CSV doesn't break the analysis
                String noCommas=doc.replace(",","");
                if(student.get(key).containsKey(doc)) {
                    //create each student's word count string
                    heading.append(noCommas).append(": ").append(student.get(key).get(doc).getWordCount()).append(" words").append(",");
                    finalCount += student.get(key).get(doc).getWordCount();
                }else{
                    heading.append(noCommas).append(": 0 words,");
                }
            }
            StringBuilder finalAmount= new StringBuilder();
            finalAmount.append(key).append(",").append("Total word count: ").append(finalCount).append(",").append(heading);
            csvWriter.flush();
            csvWriter.write(String.valueOf(finalAmount));
            csvWriter.newLine();
            heading=new StringBuilder();
            finalAmount=new StringBuilder();
            finalCount=0;
            //same process as the word count, but now for the comment count
            for(String doc:allFileNames){
                String noCommas=doc.replace(",","");
                if(student.get(key).containsKey(doc)) {
                    heading.append(noCommas).append(": ").append(student.get(key).get(doc).getContributionCount()).append(" comments").append(",");
                    finalCount += student.get(key).get(doc).getContributionCount();
                }else{
                    heading.append(noCommas).append(": 0 comments,");
                }
            }
            finalAmount.append(" ").append(",").append("Total comment count: ").append(finalCount).append(",").append(heading);
            csvWriter.flush();
            csvWriter.write(String.valueOf(finalAmount));
            csvWriter.newLine();
            csvWriter.newLine();
            heading=new StringBuilder();
            finalCount=0;
        }
        csvWriter.close();
        System.out.println("Done. Analysis File Saved in "+source+"Analysis");
    }
}

