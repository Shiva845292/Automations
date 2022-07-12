import os
from datetime import datetime
def finding_media_file(location):
    os.chdir(location)
    sorted_folder=sorted(os.listdir(),key=os.path.getmtime,reverse=True)
    for count,UPN in enumerate(sorted_folder):
        if count==49: break
        else:
            UPN_FOLDER = os.path.join(location,UPN)
            os.chdir(UPN_FOLDER)
            for folder in os.listdir():
                if folder.lower()=="media":
                    media_folder=os.path.join(UPN_FOLDER,folder)
                    os.chdir(media_folder)
                    media_file=os.listdir()

            IN_dir = "C:\\In_folder"
            os.chdir(IN_dir)
            for file in sorted(os.listdir(),key=os.path.getmtime,reverse=True):
                for i in media_file:
                    if i==file:
                        os.remove(file)
                        print(f"THE -----{file}----- IS DELETED...! FROM {UPN}")

if __name__ == "__main__":
    finding_media_file("C:\\New folder")
    # finding_media_file("C:\\pkg_ctx")


