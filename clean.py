import os, winreg, shutil, subprocess


OS = ''

def detect_os():
	global OS
	if os.name == 'nt':
		OS = 'windows'
	else:
		OS = 'unix'


def get_all_drives():
    drives = []
    for letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
        if os.path.exists(f"{letter}:\\"):
            drives.append(f"{letter}:\\")
    return drives


def empty_recycle_bin():
    drives = get_all_drives()

    for drive in drives:
        try:
            recycle_bin_path = f"{drive}$Recycle.Bin"
            if os.path.exists(recycle_bin_path):
                subprocess.run(f"rd /s /q {recycle_bin_path}", shell=True, check=True)
                print(f"Cleaned trash on {drive}")
            else:
                print(f"No recycle bin found on {drive}")
        except subprocess.CalledProcessError as e:
            print(f"Error when emptying the trash on {drive}: {e}")


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
			subprocess.check_output("del " + os.path.join(os.getenv('LOCALAPPDATA') + "\\Microsoft\\Office\\UnsavedFiles", file) + " /F", shell=True)
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
	subprocess.check_output("cleanmgr /d C: /verylowdisk", shell=True)


def clean_recent_files():
	subprocess.check_output("del /F /Q %APPDATA%\\Microsoft\\Windows\\Recent\\*", shell=True)
	subprocess.check_output("del /F /Q %APPDATA%\\Microsoft\\Windows\\Recent\\AutomaticDestinations\\*", shell=True)
	subprocess.check_output("del /F /Q %APPDATA%\\Microsoft\\Windows\\Recent\\CustomDestinations\\*", shell=True)
	subprocess.check_output("taskkill /f /im explorer.exe", shell=True)
	subprocess.check_output("start explorer.exe", shell=True)


def main():
	detect_os()
	if OS == 'windows':
		clean_reg()
		remove_temp_files()
		remove_windows_history()
		clean_recent_files()
		empty_recycle_bin()
		clean_drive_c()
	else:
		print("Not implemented yet")
	input("Press enter to close the window. >")


if __name__ == "__main__":
	main()
