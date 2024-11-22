from sharepoint.sharepoint import Sharepoint
import pprint as pp

sp = Sharepoint()

site_groups = sp.list_site_groups()

pp.pprint(site_groups)
