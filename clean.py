import os, winreg, shutil
from subprocess import check_output


OS = ''

def detect_os():
	global OS
	if os.name == 'nt':
		OS = 'windows'
	else:
		OS = 'unix'


def delete_sub_key(root, sub):
	try:
		open_key = winreg.OpenKey(root, sub, 0, winreg.KEY_ALL_ACCESS)
		num, _, _ = winreg.QueryInfoKey(open_key)
		for i in range(num):
			child = winreg.EnumKey(open_key, 0)
			delete_sub_key(open_key, child)
		try:
		   winreg.DeleteKey(open_key, '')
		except Exception:
		   pass
		finally:
		   winreg.CloseKey(open_key)
	except Exception:
		pass


def clean_reg():
	for name in ('Word', 'Outlook', 'Access', 'Excel', 'PowerPoint', 'Visio'):
		for i in range(0, 25, 1):
			path = r'SOFTWARE\\Microsoft\\Office\\' + str(i) + '.0\\' + name + '\\'
			try:
				with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, access=winreg.KEY_ALL_ACCESS) as reg:
						try:
							winreg.DeleteKey(reg, 'Place MRU')
						except:
							pass
						try:
							winreg.DeleteKey(reg, 'File MRU')
						except:
							pass
			except Exception:
				continue
			try:
				with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path + 'Reading Locations', access=winreg.KEY_ALL_ACCESS) as reg:
					count_subkeys_redaing_location = winreg.QueryInfoKey(reg)
					for count in range(count_subkeys_redaing_location[0]):
						name_of_subkey = winreg.EnumKey(reg, count)
						try:
							winreg.DeleteKey(reg, name_of_subkey)
						except:
							continue
			except Exception:
				continue

			try:
				delete_sub_key(winreg.HKEY_CURRENT_USER, path + 'User MRU')
			except Exception:
				continue
			
			try:
				with winreg.OpenKey(winreg.HKEY_USERS, None, access=winreg.KEY_ALL_ACCESS) as reg:
					count_subkeys_redaing_location = winreg.QueryInfoKey(reg)
					for count in range(count_subkeys_redaing_location[0]):
						name_of_subkey = winreg.EnumKey(reg, count)
						delete_sub_key(winreg.HKEY_USERS, name_of_subkey + 'User MRU')
			except Exception:
				continue

	print("Clean register for Office OK")


def remove_temp_files():
	try:
		for file in os.listdir(os.getenv('LOCALAPPDATA') + "\\Microsoft\\Office\\UnsavedFiles"):
			check_output("del " + os.path.join(os.getenv('LOCALAPPDATA') + "\\Microsoft\\Office\\UnsavedFiles", file) + " /F", shell=True)
		shutil.rmtree(os.getenv('LOCALAPPDATA') + "\\Microsoft\\Office\\UnsavedFiles")
	except:
		pass

	try:	
		for file in os.listdir(os.getenv('APPDATA') + "\\Microsoft\\Windows\\Recent\\AutomaticDestinations"):
			os.remove(os.path.join(os.getenv('APPDATA') + "\\Microsoft\\Windows\\Recent\\AutomaticDestinations", file))
	except:
		pass

	print("Remove from UnsavedFiles and AutomaticDestinations OK")


def remove_windows_history():
	try:
		shutil.rmtree(os.getenv('LOCALAPPDATA') + "\\Microsoft\\Windows\\History")
	except:
		pass

	print("Remove windows history OK")


def clean_drive_c():
	check_output("cleanmgr /d C: /verylowdisk", shell=True)


def clean_recent_files():
	check_output("del /F /Q %APPDATA%\\Microsoft\\Windows\\Recent\\*", shell=True)
	check_output("del /F /Q %APPDATA%\\Microsoft\\Windows\\Recent\\AutomaticDestinations\\*", shell=True)
	check_output("del /F /Q %APPDATA%\\Microsoft\\Windows\\Recent\\CustomDestinations\\*", shell=True)
	check_output("taskkill /f /im explorer.exe", shell=True)
	check_output("start explorer.exe", shell=True)


def main():
	detect_os()
	if OS == 'windows':
		clean_reg()
		remove_temp_files()
		remove_windows_history()
		clean_recent_files()
		clean_drive_c()
	else:
		print("Not implemented yet")
	input("Press enter to close the window. >")


if __name__ == "__main__":
	main()
