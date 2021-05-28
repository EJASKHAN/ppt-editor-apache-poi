package com.certificate;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.*;
import java.time.LocalDate;
import java.util.HashMap;

/**
 * @author Ejaskhan
 **/
public class CertificateCreatorMain {
    static HashMap<String, String> mappings = getFileMappings();
    public static void main(String[] args) {
        FileInputStream inputStream = null;
        try {

            LocalDate currentDate = LocalDate.now();

            String monthStr = currentDate.getMonth().toString();
            StringBuilder monthAndYear = new StringBuilder();
            monthAndYear.append(monthStr.substring(0, 1).toUpperCase());
            monthAndYear.append(monthStr.substring(1, monthStr.length()).toLowerCase());
            monthAndYear.append("-");
            monthAndYear.append(String.valueOf(currentDate.getYear()));


            inputStream = new FileInputStream("C:\\CertificateCreator\\awardees.txt");
            //test
            int success = 0;
            int count = 0;
            try (BufferedReader br
                         = new BufferedReader(new InputStreamReader(inputStream))) {
                String line;
                while ((line = br.readLine()) != null) {
                    String[] lineArray = line.split("-");
                    if(lineArray.length!=2)
                    {
                        System.out.println(" Oh, Sorry, this is a wrong input");
                        throw new RuntimeException("wrong input");
                    }
                    boolean isSuccess = createCertificate(mappings.get(lineArray[0].trim()),
                            lineArray[1].trim(),
                            monthAndYear.toString());
                    count++;
                    success = isSuccess ? success + 1 : success;
                    if (!isSuccess) {
                        System.out.println("Certificate creation failed for " + lineArray[1] + ". Please check awardees.txt");
                    }
                }
            }
            System.out.println();
            System.out.println(success + " out of " + count + " certificates created successfully.......");


        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                inputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    static boolean createCertificate(String template, String name, String monthAndYear) {
        boolean isSuccess = false;
        try {
            System.out.println(" ************************************************************************************");
            System.out.println(" Processing the certificate for " + name);
           // FileInputStream is = new FileInputStream("src/main/resources/" + template);
            InputStream is = CertificateCreatorMain.class.getResourceAsStream("/" + template);

            ZipSecureFile.setMinInflateRatio(0);
            try (XMLSlideShow ppt = new XMLSlideShow(is)) {
                is.close();


                // size of the canvas in points
                Dimension pageSize = ppt.getPageSize();
                //System.out.println("pageSize: " + pageSize);

                for (XSLFSlide slide : ppt.getSlides()) {
                    int i = 0;
                    for (XSLFShape shape : slide) {
                        int k = 0;
                        if (shape instanceof XSLFTextShape) {
                            XSLFTextShape txShape = (XSLFTextShape) shape;
                            //System.out.println("Text: i" + i + ":k" + k + " " + txShape.getText());
                            if (txShape.getText().equalsIgnoreCase("NAME")) {
                                //System.out.println("Replacing Name with custom value");
                                //txShape.setText("GOPIKA NAIR").setFontSize(36.0);
                                XSLFTextRun run = txShape.setText(name.toUpperCase());
                                run.setBold(true);
                                run.setFontSize(36.0);
                                run.setFontFamily("Arial");
                                if(template.equalsIgnoreCase(mappings.get("1"))) {
                                    run.setFontColor(Color.WHITE);
                                }


                            } else if (txShape.getText().equalsIgnoreCase("MONTH")) {
                                //System.out.println("Replacing Name with custom value");
                                XSLFTextRun run2 = txShape.setText(monthAndYear);
                                run2.setBold(true);
                                run2.setFontSize(14.0);
                                run2.setFontFamily("Arial");
                                if(template.equalsIgnoreCase(mappings.get("1"))) {
                                    run2.setFontColor(Color.WHITE);
                                }

                            }
                        } else if (shape instanceof XSLFPictureShape) {
                            XSLFPictureShape pShape = (XSLFPictureShape) shape;
                            XSLFPictureData pData = pShape.getPictureData();
                            // System.out.println("Image: i" + i + ":k" + k + " " + pData.getFileName());
                        } else {
                            System.out.println("Process me: " + shape.getClass());
                        }
                        k++;
                    }
                    i++;
                }
                System.out.println(" Creating the certificate.............");
                String fileName = name + "-"
                        + template.split("-")[0] + "(" + monthAndYear+")" + ".pptx";
                try (FileOutputStream out = new FileOutputStream("C:\\CertificateCreator\\" + fileName)) {
                    ppt.write(out);
                }
                isSuccess = true;
                System.out.println(" Done!!! File creation, file name: " + fileName);
                System.out.println(" ************************************************************************************");
            }

        } catch (Exception exception) {
            exception.printStackTrace();
        }

        return isSuccess;
    }


    static HashMap<String, String> getFileMappings() {
        HashMap<String, String> fileMap = new HashMap<>();
        fileMap.put("1", "StarOfTheQuarter-1.pptx");
        fileMap.put("2", "RisingStar-2.pptx");
        fileMap.put("3", "TeamAward-3.pptx");
        fileMap.put("4", "Above&Beyond-4.pptx");
        return fileMap;
    }
}
