# The following is my full webdriver utility class

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options


class WebDriverUtility:

	@staticmethod
	def setup_driver(): 
		"""
		Sets up the Chrome WebDriver.
		return: A configured WebDriver instance.
		"""
		chrome_options = Options()
		chrome_options.add_argument("--no-sandbox")  # Bypass OS security model
		chrome_options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems
		chrome_options.add_experimental_option('prefs', {'loggingPrefs': {'browser': 'ALL'}})
		chrome_options.binary_location = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

		# Initialize the ChromeDriver using webdriver-manager to handle the driver setup
		service = Service(executable_path="/usr/local/bin/chromedriver")

		try: 
			driver = webdriver.Chrome(service=service, options=chrome_options)
		except Exception as e: 
			print("Error starting ChromeDriver:", e)
			return None 

		return driver


