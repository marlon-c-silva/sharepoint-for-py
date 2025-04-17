# SharePoint Integration Project

This project provides a Python-based solution for interacting with SharePoint sites. It leverages the `office365` library to perform various operations such as connecting to SharePoint, retrieving list items, downloading files, and managing site groups.

## Features

- **Authentication**: Connect to SharePoint using credentials stored in environment variables.
- **Retrieve List Items**: Fetch all items from a SharePoint list as JSON or a Pandas DataFrame.
- **File Management**: Download files from SharePoint and list files in a SharePoint folder.
- **Site Management**: List all sites and site groups, including group users.

## Requirements

- Python 3.8 or higher
- Required Python libraries:
  - `office365`
  - `pandas`
  - `python-dotenv`

## Installation

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd <repository-folder>
   ```

2. Install the required Python libraries:
   ```bash
   pip install -r requirements.txt
   ```

3. Create a `.env` file in the project root and add the following environment variables:
   ```env
   SP_USERNAME=<your-sharepoint-username>
   SP_PASSWORD=<your-sharepoint-password>
   SP_SITE_URL=<your-sharepoint-site-url>
   ```

## Usage

### Initialize the SharePoint Client

```python
from sharepoint import Sharepoint

sp_client = Sharepoint()
```

### Check Connection Status

```python
status = sp_client.return_connection_status()
print(status)
```

### Retrieve List Items as JSON

```python
list_title = "Your List Title"
items = sp_client.get_all_items_from_sp_list_as_json(list_title)
print(items)
```

### Retrieve List Items as a DataFrame

```python
list_title = "Your List Title"
df = sp_client.get_all_items_from_sp_list_as_dataframe(list_title)
print(df)
```

### Download a File

```python
link_to_url = "/sites/YourSite/Shared Documents/YourFile.txt"
target_file_path = "C:/Downloads"
sp_client.download_file(link_to_url, target_file_path)
```

### List Files in a Folder

```python
folder_url = "/sites/YourSite/Shared Documents/YourFolder"
files = sp_client.list_files(folder_url)
```

### Download All Files from a Folder

```python
folder_url = "/sites/YourSite/Shared Documents/YourFolder"
download_folder_path = "C:/Downloads"
sp_client.download_files(folder_url, download_folder_path)
```

### List All Sites

```python
sites = sp_client.list_all_sites()
print(sites)
```

### List Site Groups and Users

```python
groups = sp_client.list_site_groups()
print(groups)
```

## Project Structure

```
.
├── sharepoint.py       # Main SharePoint client implementation
├── .env                # Environment variables (not included in the repository)
├── requirements.txt    # Python dependencies
└── README.md           # Project documentation
```

## Notes

- Ensure that the `.env` file is properly configured with your SharePoint credentials and site URL.
- Use secure methods to store and manage credentials.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Acknowledgments

- [Office365-REST-Python-Client](https://github.com/vgrem/Office365-REST-Python-Client) for providing the SharePoint integration library.