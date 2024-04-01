from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_utility import WebDriverUtility
from bs4 import BeautifulSoup
import openpyxl
import re
import time


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
		indexed_pages_html = self.browser_navigator.get_indexed_pages_html()
		indexed_data = self.data_parser.parse_indexed_pages(indexed_pages_html)
		self.excel_manager.update_indexed_pages(indexed_data)

		# 404s Data 
		not_found_html = self.browser_navigator.get_not_found_pages()
		not_found_data = self.data_parser.parse_not_found_pages(not_found_html)
		self.excel_manager.write_404s(not_found_data)

		# Queries Last 3 Months Data 
		queries_html = self.browser_navigator.get_queries_data()
		queries_data = self.data_parser.parse_queries_data(queries_html)
		self.excel_manager.write_queries_data(queries_data)

		# Top Pages Last 3 Months Data
		top_pages_html = self.browser_navigator.get_top_pages_data()
		top_pages_data = self.data_parser.parse_top_pages_data(top_pages_html)
		self.excel_manager.write_top_pages_data(top_pages_data)

		# Total Clicks Last 3 Months Data
		total_clicks_html = self.browser_navigator.get_total_clicks_data()
		total_clicks_data = self.data_parser.parse_total_clicks_data(total_clicks_html)
		self.excel_manager.update_total_clicks_data(total_clicks_data)

		# Save the Excel workbook 
		self.excel_manager.save_workbook("Monthly SEO Metrics.xlsx")


	def close(self): 
		"""Method to close the WebDriver when done"""
		self.browser_navigator.driver.quit()



class BrowserNavigator:


	def __init__(self, driver):
		self.driver = driver 


	def navigate_to_console(self): 
		url = 'https://search.google.com/u/2/search-console/index?resource_id=sc-domain%3Aturnkeyofficespace.com'
		self.driver.get(url)
		time.sleep(5)  # Adjust sleep time as necessary for page to load 


	def get_indexed_pages(self): 
		# Navigate to the URL that contains the indexed pages info 
		self.navigate_to_console() 

		indexed_data = {"Indexed Count": None, "Last Updated": None}

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
				WebDriverWait(self.driver, 7).until(
					EC.staleness_of(top_pages_elements[0])
				)
			except (NoSuchElementException, TimeoutException): 
				break  # Stop if no pagination arrow is found 

		return page_clicks 


	def get_total_clicks(self): 
		self.driver.get(
			'https://search.google.com/u/2/search-console/performance/search-analytics'
			'?resource_id=sc-domain%3Aturnkeyofficespace.com')

		# Assure chart and page have loaded
		WebDriverWait(self.driver, 7).until(
			EC.presence_of_element_located((By.CSS_SELECTOR, ".nnLLaf.vtZz6e"))
		)

		# Find the element containing the total clicks using its unique class names
		try:
			total_clicks_element = self.driver.find_element(By.CSS_SELECTOR, "div.nnLLaf.vtZz6e")
			total_clicks_title = total_clicks_element.get_attribute('title')
			# The title attribute contains the numerical value as a string
			# Convert the title string to an integer
			total_clicks = int(total_clicks_title.replace(',', ''))  # Remove commas from thousands
		except NoSuchElementException:
			print("Total clicks element not found.")
			total_clicks = None 

			# Fetching the most recent date
			# Find all elements that contain dates and select the last one

			dates_elements = self.driver.find_elements(By.CSS_SELECTOR, "text.V67aGc > tspan")
			if dates_elements: 
				most_recent_date = dates_elements[-1].text  # Last item in list is most recent
			else: 
				print("No date elements found.")
				most_recent_date = None 

		
		return most_recent_date, total_clicks 









# report_generator = SEOReportGenerator()
# report_generator.run()
# report_generator.close()





	
