"""
BEFORE RUNNING FOR THE GOOGLE SHEETS API:
---------------
1.  If not already done, enable the Google Sheets API
    and check the quota for your project at
    https://console.developers.google.com/apis/api/sheets
2.  Install the Python client library for Google APIs by running
    `pip install --upgrade google-api-python-client`
"""

"""
If encountering errors relating to pydrive:
---------------
1.  If not already done, enable the Google Drive API
2.  Install the Python library pydrive by running
    `pip install pydrive'
"""

"""
FOR THE XML_RPC WORDPRESS API, verify you meet the following requirements:

1.  WordPress 3.4+ OR WordPress 3.0-3.3 with the XML-RPC Modernization Plugin.
    Python 2.6+ OR Python 3.x
2.  Install from PyPI using 
    'easy_install python-wordpress-xmlrpc' 
    or
    'pip install python-wordpress-xmlrpc.'
"""

import pickle
import os.path

### Prints JSON files and dictionaries in a pretty print
from pprint import pprint

### Imports for the Google Drive and the Google Sheets APIs
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient import discovery
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

# Imports for the boto3 S3 API
import logging
import boto3
from botocore.exceptions import ClientError

### Imports for the WordPress API (Comment this portion when testing the image renaming and downloading)
from wordpress_xmlrpc import Client, WordPressPost
from wordpress_xmlrpc.methods.posts import GetPosts, NewPost
from wordpress_xmlrpc.methods.users import GetUserInfo
from wordpress_xmlrpc.methods import media, posts
from wordpress_xmlrpc.compat import xmlrpc_client
from wordpress_xmlrpc import WordPressPage
from wordpress_xmlrpc.methods import taxonomies


  
### Uses OAuth to determine authorization to the Google API clients
gauth = GoogleAuth()
gauth.LocalWebserverAuth() # client_secrets.json need to be in the same directory as the script
drive = GoogleDrive(gauth)

### Stores the location of the drive folder 
# Ex: https://drive.google.com/drive/folders/***
# The URL portion after the "/folders/" is the location of the folder
driveLocation = '***'

### Creates a list of all the files in the Google Drive folder
fileList = drive.ListFile({'q': "'%s' in parents and trashed=false" % driveLocation}).GetList()

s3UseBucket = 'storage***'

### Creates the client to access WordPress (Comment this portion out to test the image renaming and downloading)
wp = Client('http://localhost:8000/xmlrpc.php', 'user', 'pass')
foo = wp.call(GetUserInfo())
print(foo)

### Retrieves the city and state (or city and country for Canada) based on the entered address. Returns the created locality.
def getCityState (address):

    # The address is delimited by ", ". This will create a list of each string delimited by ", "
    stringDict = address.split(", ")
    length = len(stringDict)

    # When the list contains only one word, just use that as the locality
    if (length == 1):
        return stringDict[0].replace(" ", "-")

    # When the list contains two words, use those two words as the locality
    elif (length == 2):
        return stringDict[0].replace(" ", "-") + "-" + stringDict[1].replace(" ", "-")

    # When the list contains more than two words, the second word and the last word are used as the locality.
    else:
        return stringDict[1].replace(" ", "-") + "-" + stringDict[len(stringDict) - 1].replace(" ", "-")


### Retrieves the credentials necessary for authorization to access the data from the APIs
def getCredentials():

    SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    credentials = None

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            credentials = pickle.load(token)
   # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            credentials = flow.run_local_server(port=0)
      # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(credentials, token)

    return discovery.build('sheets', 'v4', credentials=credentials)

### Retrieves the file location for the files to be uploaded to S3
def getFolderLocation (isLicensedImage, isWikiImage, isGoogleImage, entry, vbID):

    # The folder path that will be used to upload to S3 and generate the necessary folders
    folderLocation = "clean/" + vbID + "/"

    # If the image was from the licensed image column, need to determine if it is unsplash or pixabay
    if (isLicensedImage):
        if ("unsplash.com" in entry):
            folderLocation += "unsplash/"

        if ("pixabay.com" in entry):
            folderLocation += "pixabay/"

    elif (isWikiImage):
        if ("wikimedia.org" in entry):
            folderLocation += "wikimedia/"
        else:
            folderLocation += "other/"

    elif (isGoogleImage):
        folderLocation += "google/"

    else:
        folderLocation += "other/"

    return folderLocation


def upload_file(file_name, bucket, object_name=None):
    """Upload a file to an S3 bucket

    :param file_name: File to upload
    :param bucket: Bucket to upload to
    :param object_name: S3 object name. If not specified then file_name is used
    :return: True if file was uploaded, else False
    """

    # If S3 object_name was not specified, use file_name
    if object_name is None:
        object_name = file_name

    # Upload the file
    s3_client = boto3.client('s3',
                    aws_access_key_id="key",
                    aws_secret_access_key="secret")
    try:
        response = s3_client.upload_file(file_name, bucket, object_name)

    except ClientError as e:
        logging.error(e)
        return False
        
    return True


### Contains the code that will rename images, download the images locally, then upload the images to WordPress
def main():

    # The credentials necessary for authorization to access the data from the APIs
    service = getCredentials()

    # The ID of the spreadsheet to retrieve data from.
    # 10 Sample (1wzpE23x9c2-2g-jZB2hc1mZ9USroux-2EuW-mNIZs2g)
    # 100 Sample Data Test (1zikEZgwuvwLDDH7fsA7al5MF-NlskGK3ZEEB58ku1bc)
    # VB Data Sample - 2000 records... (1vlD6nYRAx2XyJo7iTPLihNssuT-RMT8xBxs4_tFdopk)
    spreadsheet_id = 'id***'

    # The range of data to retrieve
    ranges = ['A:AR']

    # How values should be represented in the output.
    value_render_option = 'FORMATTED_VALUE'

    # Makes the API GET request to the Google Sheet API based on the entered parameters
    request = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id, ranges=ranges, valueRenderOption=value_render_option)
    response = request.execute()

    # The index numbers of the each of the fields based on the order of the columns from the Google Sheets
    vb_IDIndex = 0
    nameIndex = 1
    addressIndex = 2
    subCategoryIndex = 4
    licensedImage1Index = 24
    wikiImage1Index = 33
    googleImage1Index = 39
    columnAS = 44

    # The flag used to determine if the current row is the row containing the headers
    headers = True

    # Row counter
    row = 1

    # Iterates through each of the rows in the spreadsheet
    # Will format each of the image file names to the following:
    # name-city-state-subcategory-image_column_location-vb_ID
    for entry in response["valueRanges"][0]["values"]:
      
        # If the current row is the header row, skip the row
        if (headers):
            headers = False
            continue
      
        # Replaces any existing delimiters in the name string to be just a space " "
        newFileName = entry[nameIndex].replace(" - ", " ")
        newFileName = newFileName.replace(", ", " ")
        newFileName = newFileName.replace(" ", "-") + "-"
        
        # Gets the Locality of the entry
        newFileName += getCityState(entry[addressIndex]) + "-"

        # Gets the Sub-Category of the entry
        newFileName += entry[subCategoryIndex].replace(" ", "-") + "-"

        # Iterates through each image URL in Columns Y:AR
        for i in range(licensedImage1Index, columnAS):

            # Trys to do the following, otherwise move on to the next column
            try:

                # If the column contains any data, then...
                if (entry[i]):

                    # Iterate through each file in the Google Drive folder location
                    for file in fileList:

                        # If the file is an image of the "image/jpeg" mimetype then...
                        if (file['mimeType'] == "image/jpeg"):

                            # Get the name of the image. Break it by using the "." as the delimiter. Get use the first string as the searchString.
                            fileName = file['title']
                            searchList = fileName.split(".")
                            searchString = searchList[0]

                            # If the image URL contains the searchString
                            if searchString in entry[i]:

                                # Determines the image_column_location adds the vb_ID and the file type to the end of the file name.
                                suffixString = ""

                                # Checks to see where the image came from (Used for S3 folder separation)
                                isLicensedImage = False
                                isWikiImage = False
                                isGoogleImage = False

                                # If the image URL was in the LicensedImages columns, add "L" and the number indicating the LicensedImage column
                                if (i >= licensedImage1Index and i < wikiImage1Index):
                                    suffixString += "L" + str(i - licensedImage1Index + 1) + "-"
                                    isLicensedImage = True
                                
                                # If the image URL was in the WikiImages columns, add "W" and the number indicating the WikiImage column
                                elif (i >= wikiImage1Index and i < googleImage1Index):
                                    suffixString += "W" + str(i - wikiImage1Index + 1) + "-"
                                    isWikiImage = True

                                # If the image URL was in the GoogleImages columns, add "G" and the number indicating the GoogleImage column
                                elif (i >= googleImage1Index and i < columnAS):
                                    suffixString += "G" + str(i - googleImage1Index + 1) + "-"
                                    isGoogleImage = True

                                # Replaces any spaces with hyphens "-"
                                vbID = entry[vb_IDIndex].replace(" ", "-")
                                suffixString += vbID

                                # Adds the file type to the end of the suffixString
                                fileType = ".jpg"

                                # The folder path that will be used to upload to S3 and generate the necessary folders
                                folderLocation = getFolderLocation(isLicensedImage, isWikiImage, isGoogleImage, entry[i], vbID)

                                # Downloads the image file locally with the newFileName
                                downloadImage = newFileName + suffixString + fileType
                                file.GetContentFile(downloadImage)
                                print(downloadImage)
                                upload_file(downloadImage, s3UseBucket, folderLocation + downloadImage)
                                # print (folderLocation + downloadImage)
                                # print(searchString)

                                # Searches for the metadata file associated with the image
                                for dataFile in fileList:

                                    # Finds JSON files
                                    if (dataFile['title'].endswith(".json") and searchString in dataFile['title']) :
                                        fileType = ".json"
                                        downloadFile = newFileName + suffixString + fileType

                                        file.GetContentFile(downloadFile)
                                        print(downloadFile)

                                        upload_file(downloadFile, s3UseBucket, folderLocation + downloadFile)
                                        # print (folderLocation + downloadFile)
                                        break
                                    
                                    # Finds XML files
                                    if (dataFile['title'].endswith(".xml") and searchString in dataFile['title']) :
                                        fileType = ".xml"
                                        downloadFile = newFileName + suffixString + fileType

                                        file.GetContentFile(downloadFile)
                                        print(downloadFile)

                                        upload_file(downloadFile, s3UseBucket, folderLocation + downloadFile)
                                        # print (folderLocation + downloadFile)
                                        break

                                # Creates the path to the image files (Comment this portion out to test the image renaming and downloading)
                                path = '/Users/bilalmohamad/Documents/VBImages/VBSheets/'
                                filePath = path + newFileName + suffixString

                                # # Creates the data associated with the image
                                data = {
                                   'name': newFileName + suffixString,
                                   'type': 'image/jpeg', #mimetype
                                }
                                
                                # Opens the image reads the image to the xml_rpc WordPress API client
                                with open(filePath, 'rb') as img:
                                      data['bits'] = xmlrpc_client.Binary(img.read())

                                # Uploads the image to WordPress
                                wp.call(media.UploadFile(data))

                                break
            except:
                continue
            
        row += 1


# Executes the program
main()