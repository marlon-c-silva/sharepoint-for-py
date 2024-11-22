from sharepoint.sharepoint import Sharepoint
import pprint as pp

# If you have the environment variables set up, you can use this line instead of the one above
sp = Sharepoint()

# If you want to get the data from a SharePoint list as Json, you can use the following code:
sp_list = sp.get_all_items_from_sp_list_as_json("FAQ")
pp.pprint(sp_list)
