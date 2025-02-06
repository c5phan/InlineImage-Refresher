# InlineImage-Refresher
Automated Inline Image Refresher for Google Docs using AppScript. It is set up to refresh every time a document opens.
To use, please see information below...

# Guidelines for Usage
1.	Current code can only replace INLINE images meaning they can not be set to wrap text or any other layout
    a.	Works for images in header, body and footer

2.	Any image that you want to auto update must have the same matching alt text to the one in the master sheet (even a blank line in the alt text will affect the code)
    a.	You must manually add the alt text to each doc image through the image options

3.	The master sheet must also hold the correct FileID of the imagine the shared google drive
4.	The refresher automatically resizes to the dimensions of the one in the doc 
    a.	Ensure that the one in the sheet is either the same dimensions or can scale down to the dimension without distortion

# Setup

1. Personalizing the Script Code and Master Sheet

Create a Master Sheet to hold your images. You will update the sheet images whenever needed and the script will use the one in the sheet.
The format of the sheet must have Column C holding the fileID of the image (assuming it is within Google Drive) and D will be the alt-text for the script to look for.
The Alt Text must be unique as it is how the script will identify which image to use and replace.
For the script, replace the placeholder sheet URL with the share link to the sheet you want it to retriever from. 
You want to also deploy the script as a library (this will give you a Script ID that will be needed later on for set up).

3. Adding the Image to the Master Sheet

Add the image you want to be refreshed over time to the master sheet. This is what you will update with the new image whenever needed.
The most important columns are the FileID and ALT Text (the script needs these two cells to update the image)
To get the FileID, take the section of the url after d/ and before the next /
The ALT text can be anything you want, just make sure it is unique/different to the other alt text in the sheet
For Image and URL, this is option but is helpful for you to easily find an image  

3. Setting up the Doc

Open the doc that you want to add the refresher to and go to Extensions > Appscript.
Then, you want to copy the code from the access file and paste it into your document's appscript.
You can adjust the arguments according to your needs (see access comments).

Now, click the + beside Libraries and a popup will show up for the Script ID.
Enter the Script ID into the slot and click Lookup. 

If you see AutomatedImageRefresher, click Add. The Library should now show up.
Now, we want to set up a trigger so it runs when the doc opens.  Click the clock symbol (Triggers) on the left.
Click Add Trigger on the bottom right. You can change the failure notification settings to immediately if you want or anything else, then click save once you're done.
 
Finally, you can name the appscript at the top and SAVE before closing


4.	Setting up the images in the doc

Add an image and adjust the size to your doc. Make sure that it is set to the inline layout (first option).
Copy the alt text from the appropriate image in the master sheet and paste it into the alt text option of the image in the doc. (Click on the image, select image options on the tool bar, then a tab should appear on the right and go to alt text. Paste the alt text there)
Do this for any image you want to refresh and once you open/refresh the doc it should update to the one in the sheet

