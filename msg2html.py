import win32com.client
import os
import re

def convert_msg_to_html(msgFolder, msgOutputFolder, msgFile):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    msgName = msgFile.split(".")[0]

    output_html_file = msgName+".html"

    try:
        msgFile = outlook.OpenSharedItem(os.path.join(msgFolder,msgFile))
        os.makedirs(os.path.join(msgOutputFolder+msgName))

        htmlBody = msgFile.HTMLBody

        searchPattern = r'src="cid:(.*?)@'
        replacePattern =  r'src="cid:(.*?)@[0-9A-Z.]*"'

        count = htmlBody.count('src="cid:')
        for image in range(count) :
            imageName = re.search(searchPattern, htmlBody).group(1)
            htmlBody = re.sub(replacePattern, f'src="{imageName}"', htmlBody, count=1)

        # Save the HTML to a file
        with open(os.path.join(msgOutputFolder+msgName,output_html_file), 'w', encoding='utf-8') as htmlFile:
            htmlFile.write(htmlBody)

        for attachment in msgFile.Attachments:
            attachment.SaveAsFile(os.path.join(msgOutputFolder+msgName, attachment.FileName))

        print(f"{output_html_file} : done")
    
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    msgFolder = os.getcwd()
    msgOutputFolder = os.getcwd()+"\\output\\"
    if not os.path.exists(msgOutputFolder):
        os.makedirs(msgOutputFolder)

    msgs = os.listdir(msgFolder)
    for msg in msgs:
        if msg.endswith(".msg"):
            convert_msg_to_html(msgFolder, msgOutputFolder, msg)

