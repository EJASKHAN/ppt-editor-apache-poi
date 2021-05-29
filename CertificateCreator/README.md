# ppt-editor-apache-poi
This repo contains the source code for creating and editing ppt files using apache POI and VBS script for converting pptx(s) to pdf(s).

# Steps

    1. Clone this repo.
    2. Create a folder in your C drive with the name, ''CertificateCreator''.
    3. Create a file with the name, awardees.txt , or any index file for formatting
    4. Copy 'script\certificateCreator.bat' to the folder, 'CertificateCreator'.
    5. Create two subfolders 'generatedFiles' and 'plugin' inside 'CertificateCreator'.
    6. Do maven packaging, mvn clean package, and copy the jar with the dependencies to 'CertificateCreator\plugin'.
    7. Copy 'script\convertPPTtoPDF.vbs' to 'CertificateCreator\plugin'.
    8. Add contents inside the file, awardees.txt, in the following format, certificateId - name.

#Author
Ejaskhan
