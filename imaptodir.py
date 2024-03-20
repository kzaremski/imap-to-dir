'''
    Konstantin Zaremski
    19 Feb. 2024

    LICENSE
        MIT License

        Copyright (c) 2024 Konstantin Zaremski

        Permission is hereby granted, free of charge, to any person obtaining a copy
        of this software and associated documentation files (the "Software"), to deal
        in the Software without restriction, including without limitation the rights
        to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
        copies of the Software, and to permit persons to whom the Software is
        furnished to do so, subject to the following conditions:

        The above copyright notice and this permission notice shall be included in all
        copies or substantial portions of the Software.

        THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
        IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
        FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
        AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
        LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
        OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
        SOFTWARE.
'''

# Import dependencies
import re
from imapclient import IMAPClient
from tqdm import tqdm
import configparser
from pathlib import Path

# Get config from config.ini
config = configparser.ConfigParser()
config.read('config.ini')
imapServer = config['DEFAULT']['server']
imapUsername = config['DEFAULT']['username']
imapPassword = config['DEFAULT']['password']

# Message class
class EmailMessage:
    def __init__(self, message):
        # Initial parse
        self.date = str(message[b'INTERNALDATE']).replace(":", "")
        self.RFC822 = str(message[b'RFC822'])[2:-1].replace("\\n","\n").replace("\\r", "")
        # Parse RFC822
        MessageRFC822Lines = self.RFC822.split("\n")
        for line in MessageRFC822Lines:
            if "To: " in line and not hasattr(self, 'recipient'):
                self.recipient = line.replace("To: ", "", 1)
            elif "From: " in line and not hasattr(self, 'sender'):
                self.sender = line.replace("From: ", "", 1)
            elif "Subject: " in line and not hasattr(self, 'subject'):
                self.subject = line.replace("Subject: ", "", 1)
            continue

# Main method
def main():
    # Connect to IMAP server
    server = IMAPClient(imapServer, use_uid=True)
    server.login(imapUsername, imapPassword)

    # Get all IMAP folders
    imapFolders = [str(folder[2]) for folder in server.list_folders()]
    imapFolders.sort()

    # Detect gmail account
    isGmailAccount = "[Gmail]" in imapFolders

    # Display all IMAP folders
    print("IMAP Folders on Server:")
    for folderName in imapFolders:
        subfolders = folderName.split("/")
        print("  " + ("  " * int(len(subfolders) - 1)) + subfolders[-1])
    print()

    # Download all messages in each folder
    for folderName in imapFolders:
        # Create output directory
        outputDirectoryPath = Path("./output/" + folderName).mkdir(parents=True, exist_ok=True)
        
        # Skip GMAIL folder
        if folderName == "[Gmail]":
            continue

        # Select current folder
        folder = server.select_folder(folderName)
        
        # Get a list of all message IDs
        messages = server.search(criteria='ALL', charset=None)
        #print(f"{str(len(messages))} messages in \"{folderName}\"")
        
        # For each message
        for message in tqdm(messages, folderName):
            # Fetch the message by ID and instantiate a new EmailMessage object from it
            message = EmailMessage(list(server.fetch(message, ['RFC822', 'INTERNALDATE']).values())[0])

            # Create a filename
            fileName = f"{message.subject if len(message.subject) < 50 else message.subject[0:50] + '...'} [{message.date}].eml"
            fileName = re.sub(r"[/\\?%*:|\"<>\x7F\x00-\x1F]", "-", fileName)
            outputFileName = f"./output/{folderName}/{fileName}"
            
            # Write into output file
            with open(outputFileName, "w") as file:
                file.writelines(message.RFC822)
                file.flush()

if __name__ == "__main__":
    main()
