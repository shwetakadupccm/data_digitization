from ftplib import FTP

ftp = FTP('192.168.1.5')
ftp.login(user='Shweta', passwd='Shweta#123')
ftp.dir()

