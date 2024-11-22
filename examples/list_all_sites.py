from sharepoint.sharepoint import Sharepoint

sp = Sharepoint()

result = sp.list_all_sites()

i = 0
for siteProps in result:
    print("({0} of {1}) {2}".format(i, len(result), siteProps.url))
    i += 1
