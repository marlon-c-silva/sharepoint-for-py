from sharepoint.sharepoint import Sharepoint

# If you have the environment variables set up, you can use this line instead of the one above
sp = Sharepoint()

# If you don't have any person fields, you can use the following code:
df = sp.get_all_items_from_sp_list_as_dataframe(list_title="FAQ")
df.to_csv("FAQ.csv", index=False)