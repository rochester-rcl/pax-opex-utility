# Chimera (pax-opex-utility)
A graphical utility to format PAX objects and OPEX metadata for ingest into Preservica as SIPs to be synced with ArchivesSpace.

## PAX OBJECT AND OPEX METADATA GENERATOR
This portion of the application allows for the creation of Preservica-compliant SIPS using the PAX data model and OPEX metadata.

### DEPENDENCIES
This Utility requires the Digitization Work Order Plugin created by NYU Libraries, information available at:
https://guides.nyu.edu/archivesspace/development
This plugin allows for the export of a spreadsheet of data from ArchivesSpace including Archival Object Numbers, Ref IDs, Titles, Dates, and more. This data is necessary to enable the sync with ArchiveSpace and Preservica. If you are hosted by Lyrasis, they can facilitate having the plugin installed.

### USING UTILITY WITH A VPN
If you use something like a network drive to store or stage your digital assets, it is highly recommended that you are on the same network as the assets, and not using a VPN, as this can dramatically slow down some of the steps of the process (particularly zipping up PAX objects and generating checksums on big files). If you can run the Utility on a computer sitting on your job site through remote desktop, or just work on site, the Utility will work much faster. This is mainly of concern for Work from Home folks.

### PROJECT FOLDER
Choose a folder in which to deposit all the digital assets as well as the Work Order spreadsheet generated by the NYU plugin. Use the Browse button to select the folder with the assets and spreadsheet. If you would like to generate a CSV file that contains each file and a generated SHA-1 checksum for the file, leave the "Generate File/Hash Manifest" checkbox ticked. This manifest file can the be used for quality control after ingest to ensure no files were corrupted in transit.

### WORK ORDER SPREADSHEET INFORMATION
Use the Browse button to select the Work Order Spreadsheet in the Project Folder referenced above. When exporting the spreadsheet from ArchivesSpace, make sure to check all the available boxes to ensure the Utility can find the correct columns. The Worksheet Name refers to the relevant tab in Excel that contains the necessary information. The default  value of: "digitization_work_order_report" is what the NYU Plugin outputs, but this can be updated if necessary. Crucially, the "Max Row" value needs to be updated so the Utility understands the bounds of the data. Enter the Row number of the last row with data in the spreadsheet. Do not open the Work Order Spreadsheet while using the Utility, as this can cause errors.

### PRESERVATION/ACCESS FILE EXTENSIONS
This section allows the user to specify which file types in the Project Folder are for preservation and which for access. These will be routed into their respective "Representation" in the Preservica data model. The fields default to ".pdf" for access and ".tif" for preservation which is the typical example use-case for why PAX/OPEX is useful (the "one-to-many" relationship). Please be aware that the file extension needs to be exact. If you have an error, check your files extensions as there is a difference between ".tif" and ".tiff" for instance.
          
### CREATE PAX OBJECTS AND OPEX METADATA
After filling in the fields in the Utility and hitting the "Create PAX Objects and OPEX Metadata" button, the Utility will start working. Each action the underlying scripts take will be printed to this output window. The final result will be a subdirectory inside the Project Folder you specified that is named "container_YYYY-MM-DD_HH-MM-SS" (referred to hereafter as the "container folder") but with the current date and time. If you opted to generate the file/checksum manifest, that file will also be in the Project Folder with the same name as the container folder. Inside the container folder will be a list of folders that all begin with "archival_object_" which correspond to item level metadata records in ArchiveSpace. Inside each of those will be a zipped PAX object, the associated OPEX metdadata for the PAX, and the associated OPEX metadata for the archival_object_ folder. The container folder can be dragged and dropped with CloudBerry/WinSCP/FileZilla/etc into the AWS Preservica Bulk Bucket. Ensure that inside the root of the Preservica Bulk Bucket you have a folder named "opex" and that you place the container folder inside it. For OPEX Incremental Ingest I generally recommend turning off the inner and outer ingest workflows in Preservica, transferring the container folder, and then turning the workflows back on, as that has resulted in the least errors for me. The ingested assets will then appear in whatever folder you have designated to have OPEX incremental ingests dump content into. After that, you can use functions into the following section to move the assets into the folder you use to link assets with ArchivesSpace. No part of this process actually triggers the initial "Link with ArchivesSpace" workflow in Preservica. That will need to happen either on whatever schedule you have set, or as a manual trigger.

## POST INGEST QUALITY OF LIFE
This Utility also has some simple helper functions and quality control capabilities to help manage ingests.  

### PRESERVICA ADMINISTRATOR CREDENTIALS
A Preservica account crentials with the role of:

**SDB_MANAGER_USER (Tenant Manger Role)**

is necessary in order to complete these actions. Username: typically the email of the account  
**Password:** the password of the account  
**Tenancy:** the prefix at the start of your Preservica URL for example in "https://uorrcl.access.preservica.com/" the tenancy is "UORRCL" (use all caps)  
**Server:** this defaults to the standard US Preservica server name, but update as needed, especially for Enterprise customers  
**Using 2FA Checkbox:** if you have two-factor authentication enabled for accounts, check this box  
**2FA Token:** enter the token key here from your 2FA setup process. If you already have 2FA set up before using this Utility, you will have to hit the Reset Secret Key in, Preservica, then hit "Forget" and then when you attempt to log back into the system you will be prompted to setup 2FA once again. During this process you can hit a "Reveal Key" option and then copy the key somewhere safe. [This page has more information and screenshots.](https://pypreservica.readthedocs.io/en/latest/intro.html#factor-authentication)

### TEST CONNECTION
In order to test to see if the credentials were input correctly, you can hit the "Test Connection" button which will print out all the information about the root folders in your Preservica instance if configured correctly.
          
### PRESERVICA FOLDER REF IDS
The Utility can help you move assets from one folder to another in Preservica so you don't have to do a lot of onerous dragging and dropping or keep selecting "Change Archival Structure." Each of these needs the Preservica Reference identifier, the UUID Preservica assigns to each Folder and Asset.  
**OPEX Folder Ref:** the folder which the OPEX incremental ingest dumps the digital assets into  
**ASpace Folder Ref:** the folder into which the folders representing archival objects are placed before the "Link Preservica to ASpace" workflow is run in Preservica  
**Trash Folder Ref:** the folder (which is not mandatory but that I assume most people have) which acts as the recylcing bin for Preservica, where folders/assets are dropped in for deletion  
Pressing either the "Move From OPEX to ASpace Link" or the "Move From Aspace Link to Trash" will simply move folders between places, saving time from manually dragging and dropping or hitting "Change Archival Structure." These are not mandatory but just try to be helpful. Calling the API in this way tends to only move a chunk of the entities you want moved, typically petering out around 100. Simply hit the button again to resume moving stuff between folders.

### QUALITY CONTROL
Use the "Browse" button to select the CSV file that was created in the "Pre-Ingest" tab and placed in your project folder. Hitting the "Quality Control" button will do a number of things. It will use the ArchivesSpace Ref ids to look up each asset in Preservica, and then pull the Preservica Reference ID and add it to the Work Order Spreadsheet in a new column. The Work Order spreadsheet now has a full crosswalk of identifiers for both ArchivesSpace and Preservica.
          
Next the Utility will loop through each of the newly stored Preservica References, and pull each filename and checksum and compare it to filenames and checksums stored in the CSV the Utility created. If no problems are found, the utility will print out the following in the output:

***QUALITY ASSURANCE PASSED***

If there is a problem, then a dictionary of filenames and checksums will be output that can be used to investigate problems.
