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
            try:
                # Fetch the message by ID
                message = server.fetch(message, ['RFC822', 'INTERNALDATE', 'ENVELOPE']).values()
                
                # Get message attributes from INTERNALDATE and the ENVELOPE object
                msgDate = str(message[b'INTERNALDATE']).replace(":", "")
                ENVELOPE = message[b'ENVELOPE']
                msgSenderName = ENVELOPE.from_[0].name.decode() if ENVELOPE.from_[0].name else ""
                msgSenderAddr = ENVELOPE.from_[0].mailbox.decode() + "@" + ENVELOPE.from_[0].host.decode() if ENVELOPE.from_[0].host else "none"
                msgRecipientName = ENVELOPE.to[0].name.decode() if ENVELOPE.to[0].name else ""
                msgRecipientAddr = ENVELOPE.to[0].mailbox.decode() + "@" + ENVELOPE.to[0].host.decode() if ENVELOPE.to[0].host else "none"
                msgSubject = ENVELOPE.subject.decode()

                # Create a filename
                fileName = f"{msgSubject} FROM {msgSenderName} ({msgSenderAddr}) TO {msgRecipientName} ({msgRecipientAddr}) [{msgDate}].eml"
                fileName = re.sub(r"[/\\?%*:|\"<>\x7F\x00-\x1F]", "-", fileName)
                outputFileName = f"./output/{folderName}/{fileName}"

                # Parse RFC822 standard message
                MessageRFC822 = str(message[b'RFC822'])[2:-1].replace("\\n","\n").replace("\\r", "")

                # Write into output file
                with open(outputFileName, "w") as file:
                    file.writelines(MessageRFC822)
                    file.flush()
            except Exception as e:
                continue

if __name__ == "__main__":
    main()
