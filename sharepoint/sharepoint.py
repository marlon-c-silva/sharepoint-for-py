import os
from dotenv import load_dotenv
import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.listitems.collection import ListItemCollection

load_dotenv()

SP_USERNAME = os.getenv("SP_USERNAME")
SP_PASSWORD = os.getenv("SP_PASSWORD")
SP_SITE_URL = os.getenv("SP_SITE_URL")

class Sharepoint:
    def __init__(
        self,
        sp_site_url=SP_SITE_URL,
        sp_username=SP_USERNAME,
        sp_password=SP_PASSWORD,
    ):
        """
        Initializes the SharePoint client with the provided site URL, username, and password.

        Args:
            sp_site_url (str): The URL of the SharePoint site. Defaults to SP_SITE_URL.
            sp_username (str): The username for SharePoint authentication. Defaults to SP_USERNAME.
            sp_password (str): The password for SharePoint authentication. Defaults to SP_PASSWORD.
        """
        self._sp_site_url = sp_site_url
        self._username = sp_username
        self._password = sp_password
        self._credentials = UserCredential(sp_username, sp_password)
        self._connection = self.connect()

    def get_site_url(self) -> str:
        """
        Retrieves the SharePoint site URL.

        Returns:
            str: The URL of the SharePoint site.
        """
        return self._sp_site_url

    def get_username(self) -> str:
        """
        Retrieve the username.

        Returns:
            str: The username.
        """
        return self._username

    def connect(self) -> ClientContext:
        """
        Establishes a connection to the SharePoint site using the provided credentials.

        Returns:
            ClientContext: The context object for interacting with the SharePoint site.
        """
        self.ctx = ClientContext(self._sp_site_url).with_credentials(self._credentials)
        return self.ctx

    def return_connection_status(self) -> str:
        """
        Returns the connection status to the SharePoint site.

        Returns:
            str: The connection status message.
        """
        try:
            if self._connection is None:
                return "Failed to connect to SharePoint site."
            else:
                return "Successfully connected to SharePoint site."

        except Exception as e:
            return f"Failed to connect to SharePoint site. Error: {str(e)}"

    def print_progress(items):
        print("Items read: {0}".format(len(items)))

    def get_all_items_from_sp_list_as_json(self, list_title) -> ListItemCollection:
        """
        Retrieves all items from a SharePoint list.

        Args:
            list_title (str): The title of the SharePoint list.

        Returns:
            list: A list of all items in the SharePoint list.
        """
        large_list = self._connection.web.lists.get_by_title(list_title)
        all_items = large_list.items.get_all().execute_query()
        return all_items.to_json()

    def get_all_items_from_sp_list_as_dataframe(
        self, list_title, person_fields=None
    ) -> pd.DataFrame:
        """
        Retrieves all items from a SharePoint list.

        Args:
            list_title (str): The title of the SharePoint list.

        Returns:
            list: A list of all items in the SharePoint list.
        """
        large_list = self._connection.web.lists.get_by_title(list_title)
        if person_fields is not None:
            select_fields = ["*"]
            for field in person_fields:
                select_fields.append(f"{field}/Title")
                select_fields.append(f"{field}/UserName")
            all_items = (
                large_list.items.get()
                .select(select_fields)
                .expand(person_fields)
                .execute_query()
            )
            df = pd.DataFrame(all_items.to_json())
            for field in person_fields:
                df[f"{field}_Name"] = df[field].apply(lambda x: x["Title"])
                df[f"{field}_Email"] = df[field].apply(lambda x: x["UserName"])
            return df

        all_items = large_list.items.get_all().execute_query()
        df = pd.DataFrame(all_items.to_json())
        return df

    def download_file(self, link_to_url, target_file_path):
        """
        Downloads a file from SharePoint to the local machine.

        Args:
            link_to_url (str): The relative path of the file in SharePoint.
            target_file_path (str): The target path for saving the downloaded file.
        """
        file_name = os.path.basename(link_to_url)
        file_path = os.path.join(target_file_path, file_name)
        with open(file_path, "wb") as local_file:
            file = (
                self._connection.web.get_file_by_server_relative_path(link_to_url)
                .download(local_file)
                .execute_query()
            )
        print(f"[Ok] file {file_name} has been downloaded into: {target_file_path}")

    def list_files(self, target_folder_url):
        """
        Lists all files in the specified SharePoint folder.

        Args:
            target_folder_url (str): The server-relative URL of the target folder.

        Returns:
            list: A list of file objects in the specified folder.

        Example:
            files = list_files('/sites/YourSite/Shared Documents/YourFolder')
            for file in files:
                print(file.properties["ServerRelativeUrl"])
        """
        root_folder = self._connection.web.get_folder_by_server_relative_path(
            target_folder_url
        )
        files = root_folder.get_files(True).execute_query()
        for f in files:
            print(f.properties["ServerRelativeUrl"])
        return files

    def download_files(self, target_folder_url, download_folder_path):
        """
        Downloads all files from the specified SharePoint folder to a local directory.

        Args:
            target_folder_url (str): The server-relative URL of the SharePoint folder to download files from.
            download_folder_path (str): The local directory path where the files will be downloaded.

        Returns:
            None
        """
        root_folder = self._connection.web.get_folder_by_server_relative_path(
            target_folder_url
        )
        files = root_folder.get_files(True).execute_query()
        for f in files:
            self.download_file(f.properties["ServerRelativeUrl"], download_folder_path)

    def list_all_sites(self):
        """
        Retrieves all sites from the SharePoint site.

        This method uses the tenant connection to get site properties from SharePoint
        by applying filters and executes the query to retrieve the result.

        Returns:
            result: The result of the executed query containing all lists from the SharePoint site.
        """
        result = self._connection.tenant.get_site_properties_from_sharepoint_by_filters(
            ""
        ).execute_query()
        return result

    def list_site_groups(self):
        """
        Retrieves a list of site groups and their users from the SharePoint site.

        Returns:
            list: A list of dictionaries where each dictionary represents a site group.
                  Each dictionary contains the group's title and a list of users.
                  Each user is represented as a dictionary with 'title' and 'email' keys.
        """
        site_groups = (
            self._connection.web.site_groups.expand(["Users"]).get().execute_query()
        )
        site_groups_json = []
        for g in site_groups:
            group = {"title": g.title, "users": []}
            for u in g.users:
                user = {"title": u.title, "email": u.email}
                group["users"].append(user)
            site_groups_json.append(group)
        return site_groups_json
