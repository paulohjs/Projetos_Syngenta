def check_installation_package(package_call, package_instalation):
    """
    Checks if a Python package is installed and installs it if not present.

    This function attempts to import a package by its import call name. If the import
    fails due to an ImportError, indicating that the package is not installed, it then
    proceeds to install the package using pip.

    Args:
        package_call (str): The import call name of the package (e.g., 'pandas' for Pandas).
        package_instalation (str): The name of the package to install (as used by pip).

    Example:
        check_installation_package('pandas', 'pandas')
    """

    import importlib
    import subprocess
    try:
        importlib.import_module(package_call)
        print(f"{package_call} is installed.")
    except ImportError:
        print(f"{package_call} is not installed. Installing...")
        subprocess.check_call(['pip', 'install', package_instalation])
        print(f"{package_call} was installed successfully.")


list_of_package = {"pandas": "pandas",
                   "smartsheet": "smartsheet-python-sdk",
                   "openpyxl": "openpyxl",
                   "Office365": "Office365-REST-Python-Client",
                   "selenium": "selenium",
                   "webdriver_manager": "webdriver_manager",
                   "requests": "requests"}

for name, call_name in list_of_package.items():
    check_installation_package(name, call_name)
