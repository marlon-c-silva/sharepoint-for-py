from sharepoint.sharepoint import Sharepoint
import pprint as pp

# Create an instance of the Sharepoint class passing the username and password as parameters
sp = Sharepoint('https://your-sharepoint-site-url', 'username', 'password')

# If you have the environment variables set up, you can use this line instead of the one above
sp = Sharepoint()

# check Sharepoint connection status
pp.pprint(sp.return_connection_status())
pp.pprint(sp.get_site_url())
pp.pprint(sp.get_username())
