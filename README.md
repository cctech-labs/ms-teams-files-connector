# MSTeamsFilesConnector
Python package to simplify file uploads and downloads of MS Teams programmetrically


## Installation
```bash
pip install MSTeamsFilesConnector
```


## Usage
Here's an example of how to use the `MSTeamsFilesConnector` class to upload and delete files.

```python
from ms_teams_files_connector import MSTeamsFilesConnector

tfc = MSTeamsFilesConnector(tenant_id='your_tenant_id', client_id='your_client_id', client_secret='your_client_secret')

teams_group_id = 'your_group_id'
teams_channel_id = 'your_channel_id'
src_file_path = '/path/to/file.txt'
dst_file_name = 'your_file_name.txt'

# dst_file_prefix should always start and end with '/'
dst_file_prefix = '/hello/world/'

# upload file to Teams channel
tfc.upload(teams_group_id, teams_channel_id, src_file_path, dst_file_name, dst_file_prefix)

# delete a file from Teams channel
file_name = 'your_file_name.txt'
file_prefix = '/hello/world/'
tfc.delete(teams_group_id, teams_channel_id, file_name, file_prefix)

# download a file from Teams channel
src_file_name = 'your_file_name.txt'
src_file_prefix = '/hello/world/'
dst_file_path = '/path/to/file.txt'
tfc.download(teams_group_id, teams_channel_id, file_name, dst_file_path, file_prefix)

```
