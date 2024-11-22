from sharepoint.sharepoint import Sharepoint

# If you have the environment variables set up, you can use this line instead of the one above
sp = Sharepoint()

sp.download_file("Projetos Teste/Calendario 2022.pdf", "./files/")
