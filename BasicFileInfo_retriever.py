# encoding: utf-8
from string import printable
from olefile import isOleFile, OleFileIO
from re import search
from os import walk, path
from sys import argv
from chardet import detect

def get_rvt_file_version(rvt_file):
    if path.exists(rvt_file): 
        if isOleFile(rvt_file):
            rvt_ole = OleFileIO(rvt_file)
            bfi = rvt_ole.openstream("BasicFileInfo")
            file_info_bytes = bfi.read()  # read the stream once
            detected_encoding = detect(file_info_bytes)['encoding']
            # print(detected_encoding)
            if detected_encoding == "ISO-8859-1" or detected_encoding == "Windows-1252": # concerns BIM360 hosted central models
                detected_encoding = "utf-16be" #big endian but does not work in all cases, switched to unicode_escape
                try:
                    bfi_text = file_info_bytes.decode(detected_encoding, 'replace')
                except:
                    bfi_text =file_info_bytes.decode("unicode_escape", errors='ignore')
                    bfi_text = ''.join(filter(lambda x: x in printable, bfi_text))
            elif detected_encoding == "ascii":
                detected_encoding = "utf-8"
                bfi_text = file_info_bytes.decode(detected_encoding, 'replace')
                bfi_text = ''.join(filter(lambda x: x in printable, bfi_text))
            else:
                bfi_text = file_info_bytes.decode(detected_encoding)
            worksharing = search(r"Worksharing: (.*)", bfi_text)
            if worksharing:
                if "CentralUsername:" in worksharing.group(1):
                    worksharing_data = "Worksharing: "
                else:
                    worksharing_data = "Worksharing: " + worksharing.group(1)
            else:
                worksharing_data = "Worksharing: "
            username = search(r"Username: (.*)", bfi_text)
            if username:
                username_data = "Username: " + username.group(1)
            else:
                username_data = "Username: "
            cmp = search(r"Central Model Path: (.*)", bfi_text)
            if cmp:
                if ".rvt" in cmp.group(1):
                    cmp_data = "Central Model Path: " + cmp.group(1)
                    cmp_filename = cmp.group(1).split("\\")[-1]
                    try:
                        cmp_filename = cmp_filename.split("/")[-1]
                    except: 
                        pass
                else:
                    cmp = search(r"Last Save Path: (.*)", bfi_text)
                    if ".rvt" in cmp.group(1):
                        cmp_data = "Central Model Path: " + cmp.group(1)
                        cmp_filename = cmp.group(1).split("\\")[-1]
                        try:
                            cmp_filename = cmp_filename.split("/")[-1]
                        except: 
                            pass
            else:
                cmp_data = "Central Model Path: "
                cmp_filename = "-"
            rvt_format = search(r"Format: (.*)", bfi_text)
            if rvt_format:
                rvt_format_data = "Format: " + rvt_format.group(1)
            else:
                rvt_format_data = "Format: "
            build = search(r"Revit Build: (.*)", bfi_text)
            if build:
                build_data = "Build Number: " + build.group(1)
            else:
                build_data = "Build Number: "
            last_save_path = search(r"Last Save Path: (.*)", bfi_text)
            if last_save_path:
                last_save_path_data = "Last Save Path: " + last_save_path.group(1)
            else:
                last_save_path_data = "Last Save Path: "
            open_workset_default = search(r"Open Workset Default: (.*)", bfi_text)
            if open_workset_default:
                open_workset_default_data = "Open Workset Default: " + open_workset_default.group(1)
            else:
                open_workset_default_data = "Open Workset Default: "
            project_spark_file = search(r"Project Spark File: (.*)", bfi_text)
            if project_spark_file:
                project_spark_file_data = "Project Spark File: " + project_spark_file.group(1)
            else:
                project_spark_file_data = "Project Spark File: "
            central_model_identity = search(r"Central Model Identity: (.*)", bfi_text)
            if central_model_identity:
                central_model_identity_data = "Central Model Identity: " + central_model_identity.group(1)
            else:
                central_model_identity_data = "Central Model Identity: "
            locale = search(r"Locale when saved: (.*)", bfi_text)
            if locale:
                locale_data = "Locale when saved: " + locale.group(1)
            else:
                locale_data = "Locale when saved: "
            all_local_changes_saved_to_central = search(r"All Local Changes Saved To Central: (.*)", bfi_text)
            if all_local_changes_saved_to_central:
                all_local_changes_saved_to_central_data = "All Local Changes Saved to Central: " + all_local_changes_saved_to_central.group(1)
            else:
                all_local_changes_saved_to_central_data = "All Local Changes Saved to Central: "
            central_model_version_number = search(r"Central model's version number corresponding to the last reload latest: (.*)", bfi_text)
            if central_model_version_number:
                central_model_version_number_data = "Central model's version number corresponding to the last reload latest: " + central_model_version_number.group(1)
            else:
                central_model_version_number_data = "Central model's version number corresponding to the last reload latest: "
            central_model_episode = search(r"Central model's episode GUID corresponding to the last reload latest: (.*)", bfi_text)
            if central_model_episode:
                central_model_episode_data = "Central model's episode GUID corresponding to the last reload latest: " + central_model_episode.group(1)
            else:
                central_model_episode_data = "Central model's episode GUID corresponding to the last reload latest: "
            unique_document_guid = search(r"Unique Document GUID: (.*)", bfi_text)
            if unique_document_guid:
                unique_document_guid_data = "Unique Document GUID: " + unique_document_guid.group(1)
            else:
                unique_document_guid_data = "Unique Document GUID: "
            unique_document_increment = search(r"Unique Document Increment: (.*)", bfi_text)
            if unique_document_increment:
                unique_document_increment_data = "Unique Document Increment: " + unique_document_increment.group(1)
            else:
                unique_document_increment_data = "Unique Document Increment: "
            model_identity = search(r"Model Identity: (.*)", bfi_text)
            if model_identity:
                model_identity_data = "Model Identity: " + model_identity.group(1)
            else:
                model_identity_data = "Model Identity: "
            issingleusercloudmodel = search(r"IsSingleUserCloudModel: (.*)", bfi_text)
            if issingleusercloudmodel:
                issingleusercloudmodel_data = "IsSingleUserCloudModel: " + issingleusercloudmodel.group(1)
            else:
                issingleusercloudmodel_data = "IsSingleUserCloudModel: "
            author = search(r"Author: (.*)", bfi_text)
            if author:
                author_data = "Author: " + author.group(1)
            else:
                author_data = "Author: "
            
            return "-"*50+ "\n" + cmp_filename + "\n" + worksharing_data + "\n" + username_data + "\n" + cmp_data + "\n" + rvt_format_data + "\n" + build_data + "\n" + last_save_path_data + "\n" + open_workset_default_data + "\n" + project_spark_file_data + "\n" + central_model_identity_data + "\n" + locale_data + "\n" + all_local_changes_saved_to_central_data + "\n" + central_model_version_number_data + "\n" + central_model_episode_data + "\n" + unique_document_guid_data + "\n" + unique_document_increment_data + "\n" + model_identity_data + "\n" + issingleusercloudmodel_data + "\n" + author_data + "\n"
        else: 
            pass 
    else: 
        pass

if __name__ == "__main__":
    data = ''
    folder = ''
    if len(argv) > 1:
        folder = argv[1]
    else:
        folder = input("!!! Make sure your folder is sync offline before lauching the tool to make things go faster\nFolder path:")
    folder = folder.replace("\\", "\\\\")
    # folder = r"G:\Disques partagés\M11\M11_210629_DIA_DEP\3-MODELS & DRAWINGS\BIM1\LIVE\01_SALLES BLANCHES" # for tests
    print("\nLooking for BasicFileInfo in {}\n{}".format(folder, "-"*50))
    count = 0
    errors = 0
    list_of_files = []
    # walk the folder and subfolders and find rvt files
    for root, dirs, files in walk(folder):
        for file in files:
            if file.endswith(".rvt"):
                rvt_file = path.join(root, file)
                # test if file exists
                if path.exists(rvt_file):
                    list_of_files.append(rvt_file)
    lgth = len(list_of_files)
    print("Found {} RVT files, I'm on it".format(lgth))
    for rvt_file in list_of_files:
        try:
            count += 1
            data = data + "\n#{} {}".format(count, '-'*50) + get_rvt_file_version(rvt_file)
            left_count = lgth-count+1
            sign = "|"
            if left_count%2!=0:
                sign = "-"
            print("{} left {}".format(left_count, sign ), end='\r')  # Display the remaining files on one line
        except Exception as e:
            errors += 1
            data = data + "*"*50+"\nError: " + rvt_file +"\n"

    data = data + "\nTotal Files Successfully Processed: " + str(count)
    data = data + "\nTotal Errors: " + str(errors)
    print(data)
    input("Press Enter to continue...")

