def pre_ingest_documentation(window, mline, init_color, update_color):
    mline.update('''PAX OBJECT AND OPEX METADATA GENERATOR\n-----------------------------------------------------------\nDEPENDENCIES\n''', append=True, text_color_for_value=init_color)
    mline.update('''This Utility requires the Digitization Work Order Plugin created by NYU Libraries, information available at:\n\nhttps://guides.nyu.edu/archivesspace/development\n\nThis plugin allows for the export of a spreadsheet of data from ArchivesSpace including Archival Object Numbers, Ref IDs, Titles, Dates, and more. This data is necessary to enable the sync with ArchiveSpace and Preservica. If you are hosted by Lyrasis, they can facilitate having the plugin installed.\n\n''', append=True, text_color_for_value=update_color)
    mline.update('USING UTILITY WITH A VPN\n', append=True, text_color_for_value=init_color)
    mline.update('''If you use something like a network drive to store or stage your digital assets, it is highly recommended that you are on the same network as the assets, and not using a VPN, as this can dramatically slow down some of the steps of the process (particularly zipping up PAX objects and generating checksums on big files). If you can run the Utility on a computer sitting on your job site through remote desktop, or just work on site, the Utility will work much faster. This is mainly of concern for Work from Home folks.\n\n''', append=True, text_color_for_value=update_color)
    mline.update('PROJECT FOLDER\n', append=True, text_color_for_value=init_color)
    mline.update('''Choose a folder in which to deposit all the digital assets as well as the Work Order spreadsheet generated by the NYU plugin. Use the Browse button to select the folder with the assets and spreadsheet. If you would like to generate a CSV file that contains each file and a generated SHA-1 checksum for the file, leave the "Generate File/Hash Maniest" checkbox ticked. This manifest file can the be used for quality control after ingest to ensure no files were corrupted in transit.\n\n''', append=True, text_color_for_value=update_color)
    mline.update('WORK ORDER SPREADSHEET INFORMATION\n', append=True, text_color_for_value=init_color)
    mline.update('''Use the Browse button to select the Work Order Spreadsheet in the Project Folder referenced above. When exporting the spreadsheet from ArchivesSpace, make sure to check all the available boxes to ensure the Utility can find the correct columns. The Worksheet Name refers to the relevant tab in Excel that contains the necessary information. The default  value of: "digitization_work_order_report" is what the NYU Plugin outputs, but this can be updated if necessary. Crucially, the "Max Row" value needs to be updated so the Utility understands the bounds of the data. Enter the Row number of the last row with data in the spreadsheet. Do not open the Work Order Spreadsheet while using the Utility, as this can cause errors.\n\n''', append=True, text_color_for_value=update_color)
    mline.update('PRESERVATION/ACCESS FILE EXTENSIONS\n', append=True, text_color_for_value=init_color)
    mline.update('''This section allows the user to specify which file types in the Project Folder are for preservation and which for access. These will be routed into their respective "Representation" in the Preservica data model. The fields default to ".pdf" for access and ".tif" for preservation which is the typical example use-case for why PAX/OPEX is useful (the "one-to-many" relationship). Please be aware that the file extension needs to be exact. If you are have an error, check your files extensions as there is a difference between ".tif" and ".tiff" for instance.\n\n''', append=True, text_color_for_value=update_color)
    mline.update('CREATE PAX OBJECTS AND OPEX METADATA\n', append=True, text_color_for_value=init_color)
    mline.update('''After filling in the fields in the Utility and hitting the "Create PAX Objects and OPEX Metadata" button, the Utility will start working. Each action the underlying scripts take will be mline.updateed to this output window. The final result will be a subdirectory inside the Project Folder you specified that is named "container_YYYY-MM-DD_HH-MM-SS" (referred to hereafter as the "container folder") but with the current date and time. If you opted to generate the file/checksum manifest, that file will also be in the Project Folder with the same name as the container folder. Inside the container folder will be a list of folders that all begin with "archival_object_" which correspond to item level metadata records in ArchiveSpace. Inside each of those will be a zipped PAX object, the associated OPEX metdadata for the PAX, and the associated OPEX metadata for the archival_object_ folder. The container folder can be dragged and dropped with CloudBerry/WinSCP/FileZilla/etc into the AWS Preservica Bulk Bucket. Ensure that inside the root of the Preservica Bulk Bucket you have a folder named "opex" and that you place the container folder inside it. For OPEX Incremental Ingest I generally recommend turning off the inner and outer ingest workflows in Preservica, transferring the container folder, and then turning the workflows back on, as that has resulted in the least errors for me.\n\n''', append=True, text_color_for_value=update_color)
    window.refresh()

def post_ingest_documentation(window, mline, init_color, update_color, summary_color):
    mline.update('''POST INGEST QUALITY OF LIFE\n-----------------------------------------------------------\n''', append=True, text_color_for_value=init_color)
    mline.update('This Utility also has some simple helper functions and quality control capabilities to help manage ingests.\n\n', append=True, text_color_for_value=update_color)
    mline.update('PRESERVICA ADMINISTRATOR CREDENTIALS\n', append=True, text_color_for_value=init_color)
    mline.update('''A Preservica account crentials with the role of:\nSDB_MANAGER_USER (Tenant Manger Role) is necessary in order to complete these actions.\n\nUsername: typically the email of the account\nPassword: the password of the account\nTenancy: the prefix at the start of your Preservica URL for example in "https://uorrcl.access.preservica.com/" the tenancy is "UORRCL" (use all caps)\nServer: this defaults to the standard US Preservica server name, but update as needed, especially for Enterprise customers\nUsing 2FA Checkbox: if you have two-factor authentication enabled for accounts, check this box\n2FA Token: enter the token key here from your 2FA setup process.\n\nIf you already have 2FA set up before using this Utility, you will have to hit the Reset Secret Key in, Preservica, then hit "Forget" and then when you attempt to log back into the system you will be prompted to setup 2FA once again. During this process you can hit a "Reveal Key" option and then copy the key somewhere safe. This page has more information and screenshots:\n\nhttps://pypreservica.readthedocs.io/en/latest/intro.html#factor-authentication\n\n''', append=True, text_color_for_value=update_color)
    mline.update('TEST CONNECTION\n', append=True, text_color_for_value=init_color)
    mline.update('''In order to test to see if the credentials were input correctly, you can hit the "Test Connection" button which will mline.update out all the information about the root folders in your Preservica instance if configured correctly.\n\n''', append=True, text_color_for_value=update_color)
    mline.update('PRESERVICA FOLDER REF IDS\n', append=True, text_color_for_value=init_color)
    mline.update('''The Utility can help you move assets from one folder to another in Preservica so you don't have to do a lot of onerous dragging and dropping or keep selecting "Change Archival Structure." Each of these needs the Preservica Reference identifier, the UUID Preservica assigns to each Folder and Asset.\nOPEX Folder Ref: the folder which the OPEX incremental ingest dumps the digital assets into\nASpace Folder Ref: the folder into which the folders representing archival objects are placed before the "Link Preservica to ASpace" workflow is run in Preservica\nTrash Folder Ref: the folder (which is not mandatory but that I assume most people have) which acts as the recylcing bin for Preservica, where folders/assets are dropped in for deletion\n\nPressing either the "Move From OPEX to ASpace Link" or the "Move From Aspace Link to Trash" will simply move folders between places, saving time from manually dragging and dropping or hitting "Change Archival Structure." These are not mandatory but just try to be helpful. Calling the API in this way tends to only move a chunk of the entities you want moved, typically petering out around 100. Simply hit the button again to resume moving stuff between folders.\n\n''', append=True, text_color_for_value=update_color)
    mline.update('QUALITY CONTROL\n', append=True, text_color_for_value=init_color)
    mline.update('''Use the "Browse" button to select the CSV file that was created in the "Pre-Ingest" tab and placed in your project folder. Hitting the "Quality Control" button will do a number of things. It will use the ArchivesSpace Ref ids to look up each asset in Preservica, and then pull the Preservica Reference ID and add it to the Work Order Spreadsheet in a new column. The Work Order spreadsheet now has a full crosswalk of identifiers for both ArchivesSpace and Preservica.\n\n''', append=True, text_color_for_value=update_color)
    mline.update('''Next the Utility will loop through each of the newly stored Preservica References, and pull each filename and checksum and compare it to filenames and checksums stored in the CSV the Utility created. If no problems are found, the utility will mline.update out the following in the output:\n\n''', append=True, text_color_for_value=update_color)
    mline.update('***QUALITY ASSURANCE PASSED***\n\n', append=True, text_color_for_value=summary_color)
    mline.update('''If there is a problem, then a dictionary of filenames and checksums will be output that can be used to investigate problems.''', append=True, text_color_for_value=update_color)
    window.refresh()
    
def utilities_documentation(window, mline, init_color, update_color):
    mline.update('''PRESERVICA UTILITY FUNCTIONS\n-----------------------------------------------------------\nREPORTS\n''', append=True, text_color_for_value=init_color)
    mline.update('''The "Output Storage Size of Given Folder" function allows the user to input the Ref ID for a given folder and get a report of how much storage the assets that it contains take up. Like many aspects of this Utility program that involve making API calls to Preservica, this process can take a fair amount of time to generate a report on hundreds or thousands of individual files. The output window will display the filename and size in bytes for each discrete file in a given folder. The summary in green at the end of the report will display the folder title, folder Ref ID, and the total size of all the files in bytes, GB, and TB.\n\n''', append=True, text_color_for_value=update_color)
    window.refresh()
    
def about(window, mline, update_color):
    mline.update('''PAX-OPEX Utility created by John Dewees
email: john.dewees@rocheter.edu
code4lib Slack username: @John Dewees
Version 1.4.0
Last Updated: 2024-02-19
------------------------------------
Chimera icons created by Freepik - Flaticon
https://www.flaticon.com/free-icons/chimera''', append=True, text_color_for_value=update_color)
    window.refresh()
