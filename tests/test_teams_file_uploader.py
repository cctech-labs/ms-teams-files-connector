import os
import unittest
import tempfile
from ms_teams_files_connector import MSTeamsFilesConnector


class TestTeamsFileUploader(unittest.TestCase):
    # This test method tests uploading a file to the root folder of a Microsoft Teams channel.
    def test_upload_to_root(self):
        # Set the destination file name for the test upload.
        dst_file_name = "test_file_uploader.py"
        
        # Call the helper method to run the test with the given destination file name.
        self.run_test(dst_file_name)

    # This test method tests uploading a file to a subfolder of a Microsoft Teams channel.
    def test_upload_to_subfolder(self):
        # Set the destination file name and prefix for the test upload.
        dst_file_name = "test_file_uploader.py"
        dst_file_prefix = "/test_upload_folder/"
        
        # Call the helper method to run the test with the given destination file name and prefix.
        self.run_test(dst_file_name, dst_file_prefix)

    # This helper method sets up the environment variables required to connect to the Microsoft Teams API and runs the test.
    # It first creates a temporary file to upload, uploads it to the specified destination folder, downloads the uploaded file,
    # compares its content to the original file content, and then deletes the uploaded file. Finally, it cleans up the temporary files.
    def run_test(self, dst_file_name, dst_file_prefix='/'):
        # Get necessary credentials from environment variables
        tenant_id = os.environ.get("TENANT_ID")
        client_id = os.environ.get("CLIENT_ID")
        client_secret = os.environ.get("CLIENT_SECRET")
        
        # Get Teams group and channel IDs from environment variables
        teams_group_id = os.environ.get("TEAMS_GROUP_ID")
        teams_channel_id = os.environ.get("TEAMS_CHANNEL_ID")
        
        # Create a temporary file to upload
        file_upload_content = b'This is a temporary file.'
        with tempfile.NamedTemporaryFile(delete=False) as tmp_file_upload:
            tmp_file_upload.write(file_upload_content)

        # Create an instance of MSTeamsFilesConnector and upload the file
        connector = MSTeamsFilesConnector(tenant_id, client_id, client_secret)
        response = connector.upload(teams_group_id, teams_channel_id, tmp_file_upload.name, dst_file_name, dst_file_prefix)
        self.assertEqual(response, True)
        
        # Create a temporary file to download
        with tempfile.NamedTemporaryFile(delete=False) as tmp_file_download:
            pass
        # Download the uploaded file and check if its content matches with the original content
        connector.download(teams_group_id, teams_channel_id, dst_file_name, tmp_file_download.name)
        with open(tmp_file_download.name, 'rb') as f:
            downloaded_file_content = f.read()
        self.assertEqual(downloaded_file_content, file_upload_content)
        
        # Delete the uploaded file if it was uploaded successfully, and delete the temporary files
        if response:
            connector.delete(teams_group_id, teams_channel_id, dst_file_name)
        os.unlink(tmp_file_upload.name)
        os.unlink(tmp_file_download.name)

        self.assertEqual(response, True)

if __name__ == '__main__':
    unittest.main()