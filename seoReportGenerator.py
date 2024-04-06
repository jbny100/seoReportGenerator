from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException
from webdriver_utility import WebDriverUtility
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook 
from openpyxl.utils import get_column_letter
import os 
import re
from urllib.parse import urlparse


class SEOReportGenerator: 

	def __init__(self):
		# Initialize the WebDriver using WebDriverUtility
		driver = WebDriverUtility.setup_driver()

		# Pass the WebDriver instance to BrowserNavigator
		self.browser_navigator = BrowserNavigator(driver)

		self.data_parser = DataParser()
		self.excel_manager = ExcelManager()

	def run(self):
		# Navigate to the initial URL of Google Search Console
		self.browser_navigator.navigate_to_console()

		"""HTML content is obtained and passed to DataParser for each type of data required"""

		# Indexed Pages Data
		indexed_pages = self.browser_navigator.get_indexed_pages()
		indexed_data = self.data_parser.parse_indexed_pages(indexed_pages)
		self.excel_manager.update_indexed_pages(indexed_data)
		self.excel_manager.copy_indexed_pages()  # copy updated Indexed Pages to Monthly_SEO_Metrics.xlsx

		# 404s Data 
		not_found_urls = self.browser_navigator.get_404_urls()
		not_found_data = self.data_parser.parse_get_404_urls(not_found_urls)
		self.excel_manager.write_404_urls(not_found_data)

		# Queries Last 3 Months Data 
		query_searches = self.browser_navigator.get_top_queries_and_clicks()
		queries_data = self.data_parser.parse_queries_data(query_searches)
		self.excel_manager.write_queries_data(queries_data)

		# Top Pages Last 3 Months Data
		top_pages = self.browser_navigator.get_top_pages_and_clicks()
		top_pages_data = self.data_parser.parse_top_pages_and_clicks(top_pages)
		self.excel_manager.write_top_pages_data(top_pages_data)

		# Total Clicks Last 3 Months Data
		total_organic_clicks = self.browser_navigator.get_total_clicks()
		total_clicks_data = self.data_parser.parse_total_clicks_data(total_organic_clicks)
		self.excel_manager.update_total_clicks_data(total_clicks_data)
		self.excel_manager.copy_total_clicks()  # copy updated Total Clicks to Monthly_SEO_Metrics.xlsx

		# Save the Excel workbook 
		self.excel_manager.save_workbook("Monthly_SEO_Metrics.xlsx")
		

	def close(self): 
		"""Method to close the WebDriver when done"""
		self.browser_navigator.driver.quit()


class BrowserNavigator:


	def __init__(self, driver):
		self.driver = driver 


	def navigate_to_console(self): 
		url = 'https://search.google.com/u/2/search-console/index?resource_id=sc-domain%3Aturnkeyofficespace.com'
		self.driver.get(url)
		# Wait for a specific element that signifies the page has loaded
		WebDriverWait(self.driver, 10).until(
			EC.presence_of_element_located((CSS_SELECTOR, ".nnLLaf.vtZz6e"))
		)


	def get_indexed_pages(self): 
		# Navigate to the URL that contains the indexed pages info 
		self.navigate_to_console() 

		indexed_data = {"Indexed Count": None, "Last Updated": None}

		# Would this return only text that would neatly go into the indexed_data dictionary?
		try: 
			# Find the element containing the number of indexed pages
			indexed_count_element = self.driver.find_element(By.CSS_SELECTOR, ".nnLLaf.vtZz6e")
			indexed_data["Indexed Count"] = indexed_count_element.get_attribute('title')

		except NoSuchElementException: 
			print("Indexed count element not found.")

		try: 
			# Locate the element that includes "Last Updated" text
			last_updated_element = self.driver.find_element(By.XPATH, "//*[contains(text(), 'Last updated:')]")
			full_text = last_updated_element.find_element(By.XPATH, "./..").text

			# Extract the date using regex
			match = re.search(r'\d{1,2}/\d{1,2}/\d{2}', full_text)
			if match: 
				indexed_data["Last Updated"] = match.group(0)

		except NoSuchElementException:
			print("Last updated element not found.")

		# Handle scenario where neither element is found
		if not indexed_data["Indexed Count"] and not indexed_data["Last Updated"]:
			print("Failed to find indexed page data.")
		else:
			print(f"Found data: {indexed_data}")

		return indexed_data 


	def get_404_urls(self): 
		try: 
			# Navigate to the page
			self.driver.get('https://search.google.com/u/2/search-console/index/drilldown?resource_id=sc-domain%3Aturnkeyofficespace.com&item_key=CAMYDSAC')

			# Attempt to find and click the "Not found (404)" link
			not_found_link = self.driver.find_element(By.CSS_SELECTOR, 
				"span[title='Not found (404)']")
			not_found_link.click()
		except (NoSuchElementException, ElementClickInterceptedException):
			# Handle the case where the link is not found or not clickable
			print("'Not found(404)' link does not exist or is not clickable.")
			return []

		all_404_urls = []

		while True: 
			# Collect 404 URLs from the current page
			urls_elements = self.driver.find_elements(By.CSS_SELECTOR, ".00Hai")
			for element in urls_elements: 
				all_404_urls.append(element.text)

			try: 
				# Navigate the pagination button using its class
				next_page_button = self.driver.find_element(By.CSS_SELECTOR, "span.DPvwYc.fnrFqd")
				# Assuming the button becomes non-clickable or hidden when on the last page

				if not next_page_button.is_displayed() or not next_page_button.is_enabled():
					break  # Exit loop if cannot paginate further
				next_page_button.click()
				# Wait for a moment to let the page load
				WebDriverWait(self.driver, 7).until(
					EC.staleness_of(urls_elements[0])
				)
			except (NoSuchElementException, TimeoutException): 
				break  # Exit loop if pagination button not found

		return all_404_urls 


	def get_top_queries_and_clicks(self): 
		self.driver.get(
			'https://search.google.com/u/2/search-console/performance/search-analytics'
			'?resource_id=sc-domain%3Aturnkeyofficespace.com&breakdown=query')

		queries_and_clicks = {} 

		while True: 
			# Wait for the query elements to load 
			WebDriverWait(self.driver, 7).until(
				EC.presence_of_element_located((By.CSS_SELECTOR, "span.PkjLuf"))
			)

			# Extract all queries and clicks
			queries = self.driver.find_elements(By.CSS_SELECTOR, "span.PkjLuf")
			# Find all elements representing clicks using the additional class name
			clicks = self.driver.find_elements(By.CSS_SELECTOR, "span.CC8hte")

			should_break = False 
			for query, click in zip(queries, clicks): 
				# Removing commas for thousands and convert to int
				click_count = int(click.text.replace(',', ''))
				if click_count == 0:
					should_break = True  # Break the while loop once 0 clicks is encountered
					break 
				queries_and_clicks[query.text] = click_count

			if should_break: 
				break 

			try: 
				# Check for the presence of the pagination button
				next_page_button = self.driver.find_element(By.CSS_SELECTOR, "span.DPvwYc.fnrFqd")
				if not next_page_button.is_displayed() or not next_page_button.is_enabled():
					break  # Break if the pagination button is not enabled or displayed
				next_page_button.click()
				WebDriverWait(self.driver, 7).until(
					EC.staleness_of(queries[0])
				)
			except (NoSuchElementException, TimeoutException):
				break  # Break if no pagination button is found

		return queries_and_clicks 


	def get_top_pages_and_clicks(self):
		"""Navigate to the url that contains Top pages stats."""
		self.driver.get(
			'https://search.google.com/u/2/search-console/performance/search-analytics'
			'?resource_id=sc-domain%3Aturnkeyofficespace.com&breakdown=page') 

		page_clicks = {}

		# Extract all top pages URLs
		top_pages_elements = self.driver.find_elements(By.CSS_SELECTOR, ".OOHai")

		# Find all elements that contain the clicks data for top pages
		clicks_elements = self.driver.find_elements(By.CSS_SELECTOR, "span.PkjLuf.CC8hte")

		for page_element, click_element in zip(top_pages_elements, clicks_elements):
			page_url = page_element.text 
			clicks = click_element.text 

			# Ensure clicks is a valid integer before adding 
			if clicks.isdigit():
				clicks_count = int(clicks)
				# Add to dictionary if clicks > 0 to exclude pages with 0 clicks.
				if clicks_count > 0: 
					page_clicks[page_url] = clicks_count 

			# Check for and handle pagination
			try: 
				next_page_button.click()
				WebDriverWait(self.driver, 10).until(
					EC.staleness_of(top_pages_elements[0])
				)
			except (NoSuchElementException, TimeoutException): 
				break  # Stop if no pagination arrow is found 

		return page_clicks 


	def get_total_clicks(self): 
		"""Returns the dictionary total_clicks_data"""
		self.driver.get(
			'https://search.google.com/u/2/search-console/performance/search-analytics'
			'?resource_id=sc-domain%3Aturnkeyofficespace.com')

		# Assure chart and page have loaded
		WebDriverWait(self.driver, 10).until(
			EC.presence_of_element_located((By.CSS_SELECTOR, ".nnLLaf.vtZz6e"))
		)

		total_clicks_data = {}

		# Find the element containing the total clicks using its unique class names
		try:
			total_clicks_element = self.driver.find_element(By.CSS_SELECTOR, "div.nnLLaf.vtZz6e")
			total_clicks_title = total_clicks_element.get_attribute('title')
			# The title attribute contains the numerical value as a string
			# Convert the title string to an integer
			total_clicks = int(total_clicks_title.replace(',', ''))  # Remove commas from thousands

			dates_elements = self.driver.find_elements(By.CSS_SELECTOR, "text.V67aGc > tspan")
			if dates_elements:
				most_recent_date = dates_elements[-1].text # Last item in list is most recent
				total_clicks_data[most_recent_date] = total_clicks 
			else:
				print("No date elements found.")
		except NoSuchElementException:
			print("Total clicks element or date not found.")
	
		return total_clicks_data 

"""
@staticmethod decorator allows you to call the directly using the class name, without creating an instance.
For example class.method(parameter1, parameter2)
Static methods are often used for utility or helper tasks that don't need to access or 
modify the state of a class instance.
"""

class DataParser:

	@staticmethod
	def is_valid_url(url): 
		try:
			result = urlparse(url)
			# Check if the url has a scheme (eg http or https) and a netloc (domain name)
			return all([result.scheme, result.netloc])
		except Exception: 
			return False

		"""
		Example usage: 
		print(is_valid_url("https://www.example.com"))  # Output: True
		print(is_valid_url("www.example.com"))  # Output: False
		print(is_valid_url("example"))  # Output: False
		"""

	@staticmethod
	def validate_urls(urls):
		# Use is_valid_url to filter and return only valid URLs
		return [url for url in urls if is_valid_url(url)]


	@staticmethod
	def parse_indexed_pages(indexed_data):
		"""Basic validation/transformation for indexed pages data.
		'data' might be a dict with 'Indexed Count' and 'Last Updated'
		"""
		if "Indexed Count" in indexed_data and isinstance(indexed_data["Indexed Count"], str):
			indexed_data["Indexed Count"] = int(indexed_data["Indexed Count"].replace(',', ''))
		if "Last Updated" in indexed_data and isinstance(indexed_data["Indexed Count"], str):
			# Placeholder for date parsing if necessary
			pass

		parsed_indexed_data = indexed_data

		return parsed_indexed_data 

	@staticmethod
	def parse_404_urls(all_404_urls):
		"""Filter and return only valid URLs"""
		valid_404_urls = [url for url in all_404_urls if DataParser.is_valid_url(url)]

		return valid_404_urls

	@staticmethod
	def parse_queries_data(queries_and_clicks): 
		"""Example: Ensure click counts are integers and filter out any invalid entries.
		Returns the new dictionary parsed_data."""
		parsed_queries_and_clicks = {}
		for query, clicks in queries_and_clicks.items():
			if isinstance(clicks, int):
				parsed_queries_and_clicks[query] = int(clicks)

		return parsed_queries_and_clicks 

	@staticmethod
	def parse_top_pages_data(page_clicks): 
		"""Assuming 'page_clicks' is a dictionary with pages as keys and clicks as values
		Validate URLs and ensure clicks are integers."""
		for page, clicks in page_clicks.items():
			if isinstance(clicks, int) and DataParser.is_valid_url(page):
				page_clicks_parsed[page] = int(clicks)
		
		return page_clicks_parsed 

	
	@staticmethod
	def parse_total_clicks_data(total_clicks_data):
		# Additional checks performed here if needed
		return total_clicks_data 


class ExcelManager: 

	def __init__(self): 
		self.base_path = os.getcwd()  # Directory to save Excel files


	def update_indexed_pages(self, parsed_indexed_data): 
		"""Updates or creates an Excel workbook for indexed page data."""
		filepath = os.path.join(self.base_path, 'Indexed_Pages.xlsx')
		self._update_workbook(filepath, parsed_indexed_data, ["Date", "Number of Indexed Pages"])


	def update_total_clicks_data(self, total_clicks_data):
		"""Updates or creates an Excel workbook for total clicks data."""
		filepath = os.path.join(self.base_path, 'Total_clicks.xlsx')
		self._update_workbook(filepath, total_clicks_data, ["Date", "Total Clicks"])


	def write_404_urls(self, valid_404_urls):
		"""Writes the 404 urls into the designated Excel workbook."""
		self._write_to_workbook("Monthly_SEO_Metrics.xlsx", "404s", valid_404_urls, ["URL"])


	def write_queries_data(self, parsed_queries_and_clicks):
		"""Writes query data into the Excel workbook."""
		self.write_to_workbook("Monthly_SEO_Metrics.xlsx", "Queries Last 3 Months", 
			parsed_queries_and_clicks, ["Top Queries", "Clicks"])


	def write_top_pages_data(self, page_clicks_parsed): 
		"""Writes top pages data into the Excel workbook."""
		self._write_to_workbook("Monthly_SEO_Metrics.xlsx", "Top Pages Last 3 Months", 
			page_clicks_parsed, ["Top Pages", "Clicks"])


	def copy_indexed_pages(self):
		"""Copies indexed pages data from the Indexed_Pages.xlsx to the 
		Monthly_SEO_Metrics.xlsx workbook."""
		self._copy_data("Indexed_Pages.xlsx", "Monthly_SEO_Metrics.xlsx", "Indexed Pages")


	def copy_total_clicks(self): 
		"""Copies total clicks data from the Total_Clicks.xlsx to the 
		Monthly_SEO_Metrics.xlsx workbook."""
		self._copy_data("Total_Clicks.xlsx", "Monthly_SEO_Metrics.xlsx", "Total Clicks Last 3 Months")


	def save_workbook(self, workbook_name): 
		"""Saves the workbook after all updates or changes have been made."""
		workbook_path = os.path.join(self.base_path, workbook_name)
		if workbook_name in os.listdir(self.base_path): 
			wb = load_workbook(workbook_path)
		else: 
			wb = Workbook()  # create a new workbook if not existing
		wb.save(workbook_path)


	def _update_workbook(self, filepath, data, headers):
		"""Helper method to update or create a new Excel workbook with data."""
		if os.path.exists(filepath): 
			wb - load_workbook(filepath)
			ws = wb.active 
		else: 
			wb = Workbook()
			ws = wb.active() 
			ws.append(headers)  # add headers if new workbook
		for date, value in data.items():
			ws.append([date, value])
		wb.save(filepath)


	def _write_to_workbook(self, workbook_name, sheet_name, data, headers): 
		"""Helper method to write data to a specific workbook and sheet.
		Handles both dictionaries and lists as isput data."""
		filepath = os.path.join(self.base_path, workbook_name)
		if not os.path.exists(filepath):
			wb = Workbook()
			ws = wb.create_sheet(title=sheet_name)
		else: 
			wb = load_workbook(filepath)
			if sheet_name in wb.sheetnames: 
				ws = wb[sheet_name]
				ws.delete_rows(2, ws.max_row + 1)  # clear existing data from row 2 on
			else: 
				ws = wb.create_sheet(title=sheet_name)

		ws.append(headers)  # add headers for new sheet 

		if isinstance(data, dict): 
			for key, value in data.items(): 
				ws.append([key, value])
		elif isinstance(data, list):
			for item in data: 
				ws.append([item])  # each item in its own row, under the first header

		wb.save(filepath) 


	def _copy_data(self, source_file, dest_file, sheet_name):
		"""Copies data from source file to destination file, to appropriate sheet."""
		src_path = os.path.join(self.base_path, src_file)
		dest_path = os.path.join(self.base_path, dest_file) 
		src_wb = load_workbook(src_path)
		src_ws = src_wb.active
		dest_wb = load_workbook(dest_path)
		if sheet_name in dest_wb.sheetnames:
			dest_ws = dest_wb[sheet_name]
			dest_wb.remove(dest_ws)
		dest_ws = dest_wb.create_sheet(title=aheet_name)
		for row in src_ws.iter_rows():
			dest_ws.append([cell.value for cell in row])
		dest_wb.save(dest_path) 


















# report_generator = SEOReportGenerator()
# report_generator.run()
# report_generator.close()





	
