import ftplib
import urllib
import os
import shutil, sys
import socket
import paramiko
import zipfile
import types
from datetime import datetime, timedelta

def RetainFileInfoFromServer():
    server = '10.133.13.77'
    username = 'desksoa'
    password = 'desksoa'
    directory = '/DeskSOA/ndg/output/'				
    filematch = '*.csv'
    port = 23
    s = True
    FileInfo = []
    FileRef  = []
    file_list = []
    try:
				ssh = paramiko.SSHClient()
				ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
				ssh.connect('10.133.13.77', username='desksoa',password='desksoa')
				sftp = ssh.open_sftp()
				#sftp.chdir('/DeskSOA/ndg/output/')
				sftp.chdir('/DeskSOA/ndg/output/')
				dirName = sftp.listdir('.')
				file_list = dirName
				download_count = 0
				
    except socket.error, e:
				print "The Socket Error %s" % e.message
				s=e.message
    for file in file_list:         
				src_file_path = "./%s" % ( file )
				#dst_file_path = "/".join( [ 'c:/Project/FTP-Out/', file]   )
				#dst_file_path = "/".join( [ Destination, file]   )
				retry_count = 0
				while True:
						try:
							FileInfo.append(sftp.stat(src_file_path))
							download_count += 1
							break
						except Exception, err:
							if retry_count == retry_threshold:
								sftp.close() 
								return 1
						else:
							retry_count +=1
							
						try:
							FileRef.append(src_file_path)
							download_count += 1
							break
						except Exception, err:
							if retry_count == retry_threshold:
								sftp.close() 
								return 1
						else:
							retry_count +=1
    sftp.close() 
	
    #return  str(FileRef), FileInfo
    return  str(FileInfo), FileRef

################################################################################
# Written by Naveed Hasan on 02/05/2014.
# Script Description:
# 1) Deletes the existing folders AND files on FTP
# 2) Lists folders on XXXX and creates the same list of folders on FTP
# 3) Transfers files(e.g. *.txt or *.csv) from local XXXX to the corresponding folders on FTP
# 4) Deletes the list of folders on local XXXX
# Note: The "notes.txt" file on local XXX isn't transferred or deleted.
# ------------------------------------------------------------------------------#
def fetch(Source,Destination):		#serverRef,userName,passWord,fileDirectory
    
		server = '10.133.13.77'
		username = 'desksoa'
		password = 'desksoa'
		directory = '/DeskSOA/ndg/output/'				#r"\\DeskSOA\ajk\output"		
		filematch = '*.csv'
		port = 23
		s = True
        #FileDict = dict()
		TimeList = []
        
		paramiko.util.log_to_file('/paramiko.log')		
        
		try:
				ssh = paramiko.SSHClient()
				ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
				ssh.connect('10.133.13.77', username='desksoa',password='desksoa')
				sftp = ssh.open_sftp()
				#sftp.chdir('/DeskSOA/ndg/output/')
				sftp.chdir(Source)
				dirName = sftp.listdir('.')
				file_list = dirName
				download_count = 0
				
		except socket.error, e:
				print "The Socket Error %s" % e.message
				s=e.message
		FileDict = types.DictType.__new__(types.DictType, (), {})
        		
		for file in file_list:         
				src_file_path = "./%s" % ( file )
				#dst_file_path = "/".join( [ 'c:/Project/FTP-Out/', file]   )
				dst_file_path = "/".join( [ Destination, file]   )
				retry_count = 0
				while True:
					try:
						sftp.get( file, dst_file_path) #sftp.get( remote file, local file )
						download_count += 1
						timestamp1  = sftp.stat(file).st_mtime
						t  = datetime.fromtimestamp(timestamp1)
                        #FileDict[file] = timestamp1                        
						TimeList.append(t)
						break
					except Exception, err:
						if retry_count == retry_threshold:
							sftp.close() 
							return 1
					else:
						retry_count +=1
		sftp.close() 
		return TimeList		#dirName   FileDict

def SendMultipleFilesToServer(files):
        for base in files:
            lastdir = base
            for root, dirs, fls in os.walk(base):
                while lastdir != os.path.commonprefix([lastdir, root]):
                    self._send_popd()
                    lastdir = os.path.split(lastdir)[0]
                self._send_pushd(root)
                lastdir = root
                self._send_files([os.path.join(root, f) for f in fls])

# def put(localfile,remotefile):
				# ssh = paramiko.SSHClient()
				# ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
				# ssh.connect('10.133.13.77', username='desksoa',password='desksoa')
				# sftp = ssh.open_sftp()
				# #sftp.chdir('/DeskSOA/ndg/input/')
				# sftp.put(localfile,remotefile)

def put(localfile,remotefile,remoteLocation):
				ssh = paramiko.SSHClient()
				ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
				ssh.connect('10.133.13.77', username='desksoa',password='desksoa')
				sftp = ssh.open_sftp()
				#sftp.chdir('/DeskSOA/ndg/input/')
				sftp.chdir(remoteLocation)
				sftp.put(localfile,remotefile)


def put_all(localpath,remotepath):
        os.chdir(os.path.split(localpath)[0])
        parent=os.path.split(localpath)[1]
        for walker in os.walk(parent):
            try:
                self.sftp.mkdir(os.path.join(remotepath,walker[0]))
            except:
                pass
            for file in walker[2]:
                self.put(os.path.join(walker[0],file),os.path.join(remotepath,walker[0],file))

def createZipFile(path, zipFileName):
    zf = zipfile.ZipFile(zipFileName, "w")
    for dirname, subdirs, files in os.walk(path):
        zf.write(dirname)
        for filename in files:
            zf.write(os.path.join(dirname, filename))
    zf.close()
				
def readZipFile(zipFileName):
	z = zipfile.ZipFile(zipFileName, "r")
	ZipFiles_filesList = []
	nameofFiles = ''
        for filename in z.namelist():
			print filename		
			bytes = z.read(filename,'r')
			nameofFiles = nameofFiles + filename
            #nameofFiles = nameofFiles + filename + ':' + bytes + '\n'
	ZipFiles_filesList = z.namelist()
	return  ZipFiles_filesList
	
def ExtractZipFileTo(zipFileName,DirRef,*ArchieveFiles):
    z = zipfile.ZipFile(zipFileName, "r")
    z.extractall(DirRef,ArchieveFiles)

#def ReturnFileTimestampDictionary(Source):
# def ReturnFileTimestampDictionary():
    # server = '10.133.13.77'
    # username = 'desksoa'
    # password = 'desksoa'
    # directory = '/DeskSOA/ndg/output/'
    # filematch = '*.csv'
    # port = 23
    # s = True
    # paramiko.util.log_to_file('/paramiko.log')
    # FileDict= {}
    # try:
        # ssh = paramiko.SSHClient()
        # ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        # ssh.connect('10.133.13.77', username='desksoa',password='desksoa')
        # sftp = ssh.open_sftp()
        # sftp.chdir(Source)
        # dirName = sftp.listdir('.')
                # #print dirName
        # file_list = dirName
    # except socket.error, e:
                # print "The Socket Error %s" % e.message
                # s=e.message
    # for file in file_list:         
                # try:
            # timestamp1  = sftp.stat(file).st_mtime
            # FileDict[file]= timestamp1
        # except:
                    # print "In except"

        # sftp.close()
    # return FileDict
	
		
if __name__ =="__main__":
    fetch()
    # TransferFileFromLocalToServer()