from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


class WebDriverUtility:

	@staticmethod
	def setup_driver(headless=False, additional_options=None): 
		"""
		Sets up the Chrome WebDriver.

		:param headless: Boolean indicating whether to run the browser in headless mode.
        :param additional_options: List of additional option strings to be added to ChromeOptions.

		:return: A configured WebDriver instance.
		"""

		chrome_options = Options()

		if headless: 
			chrome_options.add_argument('--headless')

		# Add other options provided as parameters
		if additional_options:
			for option in additional_options:
				chrome_options.add_argument(option)


		# Initialize the driver
		service = Service(ChromeDriverManager().install())
		driver = webdriver.Chrome(service=service, options=chrome_options)

		return driver 