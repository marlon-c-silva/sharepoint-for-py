from sharepoint.sharepoint import Sharepoint

# If you have the environment variables set up, you can use this line instead of the one above
sp = Sharepoint()

# If you want to get the data from person fields, you can use the following code with the parameter person_fields=["person_field_name", "person_field_name"]:
df_with_person_field = sp.get_all_items_from_sp_list_as_dataframe(list_title="FAQ", person_fields=["Author", "Editor"])

# If you don't have any person fields, you can use the following code:
df = sp.get_all_items_from_sp_list_as_dataframe(list_title="FAQ")
