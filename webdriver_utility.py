# The following is my full webdriver utility class

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

class WebDriverUtility:

	@staticmethod
	def setup_driver(download_path="/Users/jonathanbachrach/Projects/SEOReportGenerator"): 
		"""
		Sets up the Chrome WebDriver.
		return: A configured WebDriver instance.
		"""
		chrome_options = Options()
		chrome_options.add_argument("--no-sandbox")  # Bypass OS security model
		chrome_options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems
		chrome_options.add_experimental_option('prefs', {
			'download_default_directory': download_path,
			'download.prompt_for_download': False,  # Disable download prompt 
			'download.directory_upgrade': True,  # For Chrome 
			'safebrowsing.enabled': True  # Enable safe browsing 
			})
		chrome_options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

		# Specify path to ChromeDriver explicitly instead of using webdriver-manager
		service = Service(executable_path="/usr/local/bin/chromedriver")

		try: 
			driver = webdriver.Chrome(service=service, options=chrome_options)
		except Exception as e: 
			print("Error starting ChromeDriver:", e)
			return None 

		return driver


