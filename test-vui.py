#!/usr/bin/env python
# coding: utf-8

# In[1]:


import msal
from office365.graph_client import GraphClient
import os


# In[2]:


def acquire_token():
    global token
    authority_url = 'https://login.microsoftonline.com/<tenant_id>'
    app = msal.ConfidentialClientApplication(
        authority=authority_url,
        client_id='<client_id>',
        client_credential='<client_credential>'
    )
    token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    return token


# In[3]:


def list_folders_and_files(root_folder):
    drive_items = root_folder.children
    client.load(drive_items)
    client.execute_query()
    for drive_item in drive_items:
        item_type = drive_item.folder.is_server_object_null and "file" or "folder"
        # print(drive_item)
        print("Type: {0}, Name: {1}".format(item_type, drive_item.name))
        if not drive_item.folder.is_server_object_null and drive_item.folder.childCount > 0:
            try:
                list_folders_and_files(drive_item)
            except Exception as e:
                print("Something weird !")
                print(e)


# In[7]:


def download_root(local_path):
    """
    :type local_path: str
    """
    drive_items = drive.root.children.get().execute_query()
    for drive_item in drive_items:
        if not drive_item.file.is_server_object_null:  # is file?
            # download file content
            with open(os.path.join(local_path, drive_item.name), 'wb') as local_file:
                drive_item.download(local_file)
                client.execute_query()
            print("File '{0}' has been downloaded".format(local_file.name))


# In[12]:


def download_files(remote_folder, local_path):
    """
    :type remote_folder: str
    :type local_path: str
    """
    drive_items = drive.root.children.get().execute_query()
    for drive_item in drive_items:
        # print(drive_item)
        print(drive_item.name)
        if drive_item.name == remote_folder:
            folder_items = drive_item.children
            client.load(folder_items)
            client.execute_query()
            for item in folder_items:
                if not item.file.is_server_object_null:  # is file?
                    # print(item)
                    print(item.name)
                    # download file content
                    with open(os.path.join(local_path, item.name), 'wb') as local_file:
                        try:
                            item.download(local_file)
                            client.execute_query()
                        except Exception as e:
                            print("Error downloading")
                            print(e)
                        else:
                            print("File '{0}' has been downloaded".format(local_file.name))


# In[5]:


client = GraphClient(acquire_token)
drive = client.users["<user_id>"].drive


# In[6]:


# list files
list_folders_and_files(drive.root)


# In[8]:


# download files from OneDrive
local_path = "C:/Temp"

# root files
try:
    download_root(local_path)
except Exception as e:
    print("Error downloading")
    print(e)


# In[16]:


drive_folder = "testing"
try:
    download_files(drive_folder, local_path)
except Exception as e:
    print(e)

