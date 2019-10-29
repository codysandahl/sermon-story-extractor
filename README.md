# sermon-story-extractor
Extracts bolded text from sermons to help pastors determine which stories they have already told
* Parses sermons stored as DOCX and extracts the text in bold
* Extracted text is stored in a text file "sermon_stories.txt" in the same root directory as this Python file

## Requirements
* Python 3.6 or greater
* pip3 install -r requirements.txt

## Folder Structure
Assumes this Python file resides in a root directory, and sermons are stored in their own folders.
* Date for the sermon should be the first part of the folder name.
* Sermon should be called "Sermon.docx".

EXAMPLE:  
    ./Sermon+Story+Extractor.py  
    ./2019-10-27/Sermon.docx  
    ./2019-10-20/Sermon.docx  

## Usage
On Linux...
    ./run_sermon_story_extractor.sh

On Windows...
    run_sermon_story_extractor.bat

After the program has completed, the extracted text will be stored in "sermon_stories.txt". I use Evernote, so I store all of the stories in a single giant note that I can search when I want to see if I have used this story in a previous sermon.

