import json
import requests
import time

class MSTeamsFilesConnector:
    def __init__(self, tenant_id, client_id, client_secret, verbose=False) -> None:
        """
        Initializes the MSTeamsFileConnector object with the provided credentials.

        Args:
            tenant_id (str): The Azure Active Directory tenant ID.
            client_id (str): The client ID of the Azure Active Directory application.
            client_secret (str): The client secret of the Azure Active Directory application.
            verbose (bool, optional): Whether to print verbose output. Defaults to False.
        """
        self._tenant_id = tenant_id
        self._client_id = client_id
        self._client_secret = client_secret
        self._verbose = verbose
        self._access_token = None
        self._at_exp_time = 0
    
    def _generate_access_token(self):
        time_now = time.time()
        
        if time_now - (self._at_exp_time - 120) > 0:

            # Set the URL to request an access token
            token_url = f"https://login.microsoftonline.com/{self._tenant_id}/oauth2/v2.0/token"

            # Set the data for the HTTP POST request to get the access token
            data = {
                "grant_type": "client_credentials",
                "client_id": self._client_id,
                "client_secret": self._client_secret,
                "scope": "https://graph.microsoft.com/.default"
            }

            # Send the HTTP POST request to get the access token
            token_response = requests.post(token_url, data=data)

            # Parse the response and extract the access token
            response_json = token_response.json()
            self._access_token = response_json.get("access_token")
            self._at_exp_time = time_now + response_json.get("expires_in")

        # Print the access token
        if self._verbose:
            print(self._access_token)

    def upload(self, teams_group_id: str, teams_channel_id: str, src_file_path: str, dst_file_name: str, dst_file_prefix: str = "/"):
        """
        Uploads a file to a Microsoft Teams channel.

        Args:
            teams_group_id (str): The ID of the Microsoft Teams group associated with the channel.
            teams_channel_id (str): The ID of the Microsoft Teams channel to upload the file to.
            src_file_path (str): The full path of the file to upload.
            dst_file_name (str): The name of the file to be saved in the Teams channel.
            dst_file_prefix (str, optional): The prefix for the file path in the Teams channel. Defaults to "/".
        
        Returns:
            bool: True if the file was uploaded successfully, False otherwise.
        """
        try:
            self._generate_access_token()
            
            # Get the SharePoint site ID for the Teams site associated with the Teams group
            site_id = self._get_site_id(teams_group_id)

            # Get the SharePoint drive ID
            drive_id = self._get_drive_id(site_id)
            
            if drive_id is not None:
                # Get the Channel name
                teams_channel_name = self._get_teams_channel_name(teams_group_id, teams_channel_id)
                
                create_upload_session_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{teams_channel_name}{dst_file_prefix}{dst_file_name}:/createUploadSession"
                headers = {
                    "Authorization": f"Bearer {self._access_token}",
                    "Content-Type": "application/json"
                }
                payload = {
                    "item": {
                        "@microsoft.graph.conflictBehavior": "replace",
                        "name": dst_file_name
                    }
                }
                response = requests.post(create_upload_session_url, headers=headers, data=json.dumps(payload))

                # Parse the response and get the upload URL and expiration date
                if response.status_code == requests.codes.ok:
                    upload_url = response.json().get("uploadUrl")
                    expiration_date = response.json().get("expirationDateTime")
                    if self._verbose:
                        print(f"Upload session created. Expiration date: {expiration_date}")

                    # Open the file and read its contents
                    with open(src_file_path, "rb") as f:
                        file_contents = f.read()

                    # Set the HTTP headers with the content range
                    file_size = len(file_contents)
                    headers = {
                        "Content-Length": f"{file_size}",
                        "Content-Range": f"bytes 0-{file_size-1}/{file_size}"
                    }

                    # Send the HTTP PUT request to upload the file to the Teams channel
                    response = requests.put(upload_url, headers=headers, data=file_contents)

                    # Check if the request was successful and print the response status code
                    if response.status_code == requests.codes.ok or response.status_code == requests.codes.created:
                        if self._verbose:
                            print("File uploaded successfully.")
                        return True
                    
                    else:
                        if self._verbose:
                            print(f"Error uploading file: {response.text}")
                else:
                    if self._verbose:
                        print(f"Error creating upload session: {response.text}")

            return False
        except Exception as e:
            return False

    def _get_teams_channel_name(self, teams_group_id, teams_channel_id):
        url = f"https://graph.microsoft.com/beta/teams/{teams_group_id}/channels/{teams_channel_id}"
        headers = {"Authorization": f"Bearer {self._access_token}"}
        response = requests.get(url, headers=headers)
        teams_channel_name = response.json()["displayName"]
        return teams_channel_name

    def _get_drive_id(self, site_id):
        url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        headers = {"Authorization": f"Bearer {self._access_token}"}
        response = requests.get(url, headers=headers)
        drive_id = None
        for item in response.json()["value"]:
            if item["name"] == "Documents":
                drive_id = item["id"]
        return drive_id

    def _get_site_id(self, teams_group_id):
        url = f"https://graph.microsoft.com/v1.0/groups/{teams_group_id}/sites/root"
        headers = {"Authorization": f"Bearer {self._access_token}"}
        response = requests.get(url, headers=headers)
        site_id = response.json()["id"]
        return site_id

    def delete(self, teams_group_id: str, teams_channel_id: str, file_name: str, file_prefix: str = "/"):
        """
        Deletes a file from a Microsoft Teams channel.

        Args:
            teams_group_id (str): The ID of the Microsoft Teams group associated with the channel.
            teams_channel_id (str): The ID of the Microsoft Teams channel to delete the file from.
            file_name (str): The name of the file to delete.
            file_prefix (str, optional): The prefix for the file path in the Teams channel. Defaults to "/".
        
        Returns:
            bool: True if the file was deleted successfully, False otherwise.
        """
        self._generate_access_token()
        
        site_id = self._get_site_id(teams_group_id)
        drive_id = self._get_drive_id(site_id)
        
        if drive_id is not None:
            teams_channel_name = self._get_teams_channel_name(teams_group_id, teams_channel_id)
            
            file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{teams_channel_name}{file_prefix}{file_name}"
            headers = { "Authorization": f"Bearer {self._access_token}" }
            response = requests.get(file_url, headers=headers)
            file_id = response.json()['id']

            delete_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}"
            headers = { "Authorization": f"Bearer {self._access_token}" }
            response = requests.delete(delete_url, headers=headers)
            if response.status_code == requests.codes.no_content:
                if self._verbose:
                    print("File deleted successfully.")
                return True
            else:
                if self._verbose:
                    print(f"Error deleting file: {response.text}")
                return False
            
        else:
            return False        
    
    def download(self, teams_group_id: str, teams_channel_id: str, src_file_name: str, dst_file_path, src_file_prefix: str = "/"):
        """
        Downloads a file from a Microsoft Teams channel and saves it to disk.

        Args:
            teams_group_id (str): The ID of the Microsoft Teams group associated with the channel.
            teams_channel_id (str): The ID of the Microsoft Teams channel to download the file from.
            src_file_name (str): The name of the file to download.
            dst_file_path (str): The path of the file to save the downloaded file to.
            src_file_prefix (str, optional): The prefix for the file path in the Teams channel. Defaults to "/".

        Returns:
            bool: True if the file was downloaded successfully, False otherwise.
        """
        self._generate_access_token()
        
        site_id = self._get_site_id(teams_group_id)
        drive_id = self._get_drive_id(site_id)
        
        if drive_id is not None:
            teams_channel_name = self._get_teams_channel_name(teams_group_id, teams_channel_id)
            
            file_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{teams_channel_name}{src_file_prefix}{src_file_name}"
            headers = { "Authorization": f"Bearer {self._access_token}" }
            response = requests.get(file_url, headers=headers)
            file_id = response.json()['id']

            download_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{file_id}/content'
            headers = { "Authorization": f"Bearer {self._access_token}" }
            response = requests.get(download_url, headers=headers)

            # Save the downloaded file to disk
            with open(dst_file_path, 'wb') as f:
                f.write(response.content)
            return True
        else:
            return False