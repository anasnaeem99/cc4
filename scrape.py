from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    ElementNotInteractableException,
    ElementClickInterceptedException
)
from flask import Flask, request, jsonify
from flask_cors import CORS
import time

app = Flask(__name__)
CORS(app)

@app.route('/scrape', methods=['POST'])
def scrape():
    data = request.get_json()
    required_params = [
        'base_urls', 'input_ids', 'input_values', 'class_names',
        'div_class_names', 'master_Ids', 'source_datas', 'country_Names'
    ]
    if not all(param in data for param in required_params):
        return jsonify({"error": "Missing required input parameters"}), 400

    chrome_options = Options()
    #chrome_options.add_argument("--headless")  # Headless mode
    driver = webdriver.Chrome(options=chrome_options)

    all_scraped_data = []
    
    try:
        for i, base_url in enumerate(data['base_urls']):
            if data['country_Names'][i] == 'Turkey':
                    # Handle the specific case for Turkey
                    try:
                        print(f"Processing {base_url}...")
                        driver.get(base_url)
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                        # Add any additional scraping logic for Turkey here if needed
                        print("Successfully loaded Turkey's webpage.")
                    except Exception as e:
                        print(f"Error processing Turkey's URL: {e}")
                        continue
                


                # Handle specific case for Germany
                if data['country_Names'][i] == 'Germany':
                    try:
                        accept_button = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "docOutputPromptForm:acceptLink"))
                        )
                        accept_button.click()
                        time.sleep(3)  # Allow time for the page to transition
                        print("Accepted terms for Germany.")
                    except NoSuchElementException:
                        print("Accept button not found for Germany.")
                    except Exception as e:
                        print(f"Error clicking accept button for Germany: {e}")

                # Handle specific case for Romania
                if data['country_Names'][i] == 'Romania':
                    try:
                        # Click the "Search" button to open the modal
                        search_button = driver.find_element(By.CSS_SELECTOR, 'button[data-toggle="modal"][data-target="#searchModal"]')
                        search_button.click()

                        # Wait for the modal to appear
                        time.sleep(2)
                        print("Accepted terms for Romania.")
                    except NoSuchElementException:
                        print("Accept button not found for Romania.")
                    except Exception as e:
                        print(f"Error clicking accept button for Romania: {e}")
                
                # Handle specific case for Romania
                if data['country_Names'][i] == 'Chile':
                    try:
                        # Locate and click the checkbox
                        checkbox = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "ctl00_ContentPlaceHolder1_chkTipoBusqueda_1"))
                        )
                        checkbox.click()  # Click the checkbox
                        print("Checkbox clicked successfully.")
                    except Exception as e:
                        print(f"Failed to click the checkbox. Error: {e}")     

                # Language Switch
                if data['country_Names'][i] not in ['Turkey', 'Germany', 'Sweden', 'Bulgaria', 'France', 'Malta', 'United States Minor Outlying Islands (the)', 'Ghana', 'Canada', 'Netherlands (the)', 'Chile', 'Korea (the Republic of)', 'Saudi Arabia', 'Nigeria']:
                    try:
                        language_button = driver.find_element(By.ID, "langForm:langEN_not_selected")
                        if language_button.is_displayed():
                            language_button.click()
                            time.sleep(3)
                    except NoSuchElementException:
                        print("Language button not found.")
                    except Exception as e:
                        print(f"Error with language button: {e}")


                # Search Input
                if data['country_Names'][i] not in ['Turkey', 'Sweden', 'Bulgaria', 'Lithuania', 'Ghana', 'France', 'Hungary', 'Ireland']:
                    try:
                        input_id = data['input_ids'][i]
                        elem = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.ID, input_id))
                        )
                        elem.clear()
                        elem.send_keys(data['input_values'][i])
                        elem.send_keys(Keys.RETURN)
                        time.sleep(16)  # Wait for search results to load
                    except NoSuchElementException:
                        print(f"Search input id not found for {base_url}")
                        continue

                if data['country_Names'][i] not in ['Turkey', 'Greece', 'Portugal', 'Zimbabwe', 'United States of America', 'Romania', 'Slovakia', 'Germany', 'Norway', 'Malta', 'Nigeria','Sweden', 'Bulgaria', 'Ghana', 'United States Minor Outlying Islands (the)', 'Canada', 'Austria', 'Cyprus', 'Netherlands (the)', 'Latvia', 'Chile', 'Korea (the Republic of)', 'Saudi Arabia']:
                    try:
                        input_id = data['input_ids'][i]
                        elem = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.NAME, input_id))
                        )
                        elem.clear()
                        elem.send_keys(data['input_values'][i])
                        elem.send_keys(Keys.RETURN)
                        time.sleep(16)  # Wait for search results to load
                    except NoSuchElementException:
                        print(f"Search input name not found for {base_url}")
                        continue

                if data['country_Names'][i] == 'Ghana':
                    try:
                        elem = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[type="search"]'))
                        )
                        elem.clear()
                        elem.send_keys(data['input_values'][i])
                        elem.send_keys(Keys.RETURN)
                        time.sleep(16)  # Wait for search results to load
                    except NoSuchElementException:
                        print(f"Search input name not found for {base_url}")
                        continue
                
                if data['country_Names'][i] == 'Sweden':
                    try:
                        # Wait for the results to load (you can adjust the sleep time if needed)
                        time.sleep(6)

                    except Exception as e:
                        print(f"Search field not found by ID. Error: {e}")

                # Scrape Table
            def scrape_table():
                if data['country_Names'][i] in ['Finland', 'Cyprus', 'Zimbabwe', 'Germany', 'Greece', 'United States Minor Outlying Islands (the)', 'Netherlands (the)', 'Latvia', 'Korea (the Republic of)', 'Saudi Arabia', 'Ireland', 'Austria', 'Canada', 'Nigeria', 'Norway', 'Ghana', 'Portugal', 'Slovakia']:
                    rows = driver.find_elements(By.XPATH, f'//table[@class="{data["class_names"][i]}"]/tbody/tr')
                    if not rows:
                        try:
                            div_elem = driver.find_element(By.XPATH, f'//div[@class="{data["div_class_names"][i]}"]')
                            rows = div_elem.find_elements(By.XPATH, './/table/tbody/tr')
                        except NoSuchElementException:
                            print("Div with table not found.")
                            return []

                    page_data = []
                    for row in rows:
                        columns = row.find_elements(By.TAG_NAME, "td")
                        row_data = []

                        for column in columns:
                            td_text = column.text.strip()
                            if not td_text:
                                span_elements = column.find_elements(By.TAG_NAME, "span")
                                td_text = span_elements[0].text.strip() if span_elements else ""
                            if not td_text:
                                font_elements = column.find_elements(By.TAG_NAME, "font")
                                td_text = font_elements[0].text.strip() if font_elements else ""
                            row_data.append(td_text)

                        if any(row_data):
                            page_data.append({
                                'data': row_data,
                                'master_Ids': data['master_Ids'][i],
                                'source_datas': data['source_datas'][i]
                            })
                    return page_data

                    

                if data['country_Names'][i] == 'United States of America':
                    rows = driver.find_elements(By.XPATH, f'//table[@class="{data["class_names"][i]}"]/tbody/tr')
                    if not rows:
                        try:
                            div_elem = driver.find_element(By.XPATH, f'//div[@class="{data["div_class_names"][i]}"]')
                            rows = div_elem.find_elements(By.XPATH, './/table/tbody/tr')
                        except NoSuchElementException:
                            print("Div with table not found.")
                            return []

                    page_dataAmerica = []
                    for row in rows:
                        columns = row.find_elements(By.TAG_NAME, "td")
                        row_data = []

                        for column in columns:
                            # Default scraping logic for other countries
                            td_text = column.text.strip()  # Get text from the <td> element
                            if td_text:  # If text is found in the <td>, use it
                                row_data.append(td_text)
                            else:
                                # If no text in <td>, look for a <span> inside the <td>
                                spans = column.find_elements(By.TAG_NAME, "span")  # Find all <span> elements inside <td>
                                if spans:  # Check if any <span> exists
                                    span_text = spans[0].text.strip()  # Take the text from the first <span> element
                                else:
                                    span_text = "No text"  # Fallback text if no <span> is found
                                row_data.append(span_text)

                        if any(row_data):
                            page_dataAmerica.append({
                                'data': row_data,
                                'master_Ids': data['master_Ids'][i],
                                'source_datas': data['source_datas'][i]
                            })
                    return page_dataAmerica

                if data['country_Names'][i] == 'Lithuania':
                    all_page_dataLithuania = []
    
                    try:
                        # Find all result items
                        result_items = driver.find_elements(By.CLASS_NAME, "result-item")
                        
                        # If no result items are found, print a message and return empty data
                        if not result_items:
                            print("No results found.")
                            return all_page_dataLithuania
                        
                        # Iterate over each result item
                        for result in result_items:
                            row_data = []
                            
                            # Extract the drug name from <p class="results-header"> and <span style="color:green">
                            drug_name_element = result.find_element(By.XPATH, './/p[@class="results-header"]/span')
                            drug_name = drug_name_element.text.strip() if drug_name_element else "No drug name"
                            row_data.append(drug_name)
                            
                            # Extract the active substance from <span> (text "Active substance(s):")
                            active_substance = result.find_element(By.XPATH, './/span[text()="Active substance(s):"]/following-sibling::span').text.strip() if result.find_elements(By.XPATH, './/span[text()="Active substance(s):"]/following-sibling::span') else "No active substance"
                            row_data.append(active_substance)
                            
                            # Extract strength from <span> (text "Strength:")
                            strength = result.find_element(By.XPATH, './/span[text()="Strength:"]/following-sibling::span').text.strip() if result.find_elements(By.XPATH, './/span[text()="Strength:"]/following-sibling::span') else "No strength"
                            row_data.append(strength)
                            
                            # Extract form from <span> (text "Form:")
                            form = result.find_element(By.XPATH, './/span[text()="Form:"]/following-sibling::span').text.strip() if result.find_elements(By.XPATH, './/span[text()="Form:"]/following-sibling::span') else "No form"
                            row_data.append(form)
                            
                            # Extract supply status from <span> (text "Supply status:")
                            supply_status = result.find_element(By.XPATH, './/span[text()="Supply status:"]/following-sibling::span').text.strip() if result.find_elements(By.XPATH, './/span[text()="Supply status:"]/following-sibling::span') else "No supply status"
                            row_data.append(supply_status)

                            # Add 'master_Ids' and 'source_datas' to the row
                            master_id = data['master_Ids'][i] if 'master_Ids' in data else "No master ID"
                            source_data = data['source_datas'][i] if 'source_datas' in data else "No source data"
                            
                            # Combine all data into a dictionary
                            all_page_dataLithuania.append({
                                'data': row_data,
                                'master_Ids': master_id,
                                'source_datas': source_data
                            })

                    except Exception as e:
                        print(f"Error while scraping the result items: {e}")
                    
                    return all_page_dataLithuania

                if data['country_Names'][i] == 'Romania':
                    all_page_dataRomania = []

                    try:
                        # Find the rows in the table using XPath, which should be inside <tbody> of the table
                        rows = driver.find_elements(By.XPATH, '//table[@class="table table-striped table-condensed table-bordered"]/tbody/tr')
                        
                        if not rows:
                            print("No rows found in the table.")
                            return all_page_dataRomania
                        
                        # Iterate through each row to extract the data
                        for row in rows:
                            columns = row.find_elements(By.TAG_NAME, "td")  # Get all <td> elements in the row
                            row_data = []
                            
                            for column in columns:
                                td_text = column.text.strip()  # Get text from <td>
                                
                                if td_text:
                                    row_data.append(td_text)
                                else:
                                    # If text is missing, look for <span> inside the <td>
                                    span = column.find_element(By.TAG_NAME, "span") if column.find_elements(By.TAG_NAME, "span") else None
                                    span_text = span.text.strip() if span else "No text"
                                    row_data.append(span_text)

                            # Add 'master_Ids' and 'source_datas' to the row
                            master_id = data['master_Ids'][i] if 'master_Ids' in data else "No master ID"
                            source_data = data['source_datas'][i] if 'source_datas' in data else "No source data"

                            # Combine all data into a dictionary
                            all_page_dataRomania.append({
                                'data': row_data,
                                'master_Ids': master_id,
                                'source_datas': source_data
                            })
                                
                    except Exception as e:
                        print(f"Error while scraping the table: {e}")
                    
                    return all_page_dataRomania

                if data['country_Names'][i] == 'Hungary':
                    all_page_dataHungary = []
                    try:
                        # Find all rows with the class 'table__line line'
                        rows = driver.find_elements(By.XPATH, '//div[contains(@class, "table__line line")]')
                        
                        for row in rows:
                            # Extract the text from each 'cell' inside the row
                            columns = row.find_elements(By.CLASS_NAME, "cell")
                            row_data = [column.text.strip() for column in columns]

                            # Add 'master_Ids' and 'source_datas' to the row
                            master_id = data['master_Ids'][i] if 'master_Ids' in data else "No master ID"
                            source_data = data['source_datas'][i] if 'source_datas' in data else "No source data"

                            # Combine all data into a dictionary
                            all_page_dataHungary.append({
                                'data': row_data,
                                'master_Ids': master_id,
                                'source_datas': source_data
                            })

                    except Exception as e:
                        print(f"Error scraping data from div structure: {e}")
                    return all_page_dataHungary

                if data['country_Names'][i] == 'Sweden':
                    all_page_dataSweden = []
                    
                    try:
                        # Locate the div with the class 'medprod-table-result ng-tns-c22-10 ng-star-inserted'
                        div_elem = driver.find_element(By.CLASS_NAME, "medprod-table-result.ng-tns-c22-10.ng-star-inserted")

                        # Find all <tr> rows inside the div
                        rows = div_elem.find_elements(By.TAG_NAME, "tr")

                        # Iterate through each row and extract the text data
                        for row in rows:
                            # Extract each <td> data from the row
                            columns = row.find_elements(By.TAG_NAME, "td")
                            row_data = [column.text.strip() for column in columns if column.text.strip()]  # Directly extract text

                            # Add 'master_Ids' and 'source_datas' to the row
                            master_id = data['master_Ids'][i] if 'master_Ids' in data else "No master ID"
                            source_data = data['source_datas'][i] if 'source_datas' in data else "No source data"

                            # Combine all data into a dictionary
                            all_page_dataSweden.append({
                                'data': row_data,
                                'master_Ids': master_id,
                                'source_datas': source_data
                            })

                    except Exception as e:
                        print(f"Error scraping table: {e}")
                    return all_page_dataSweden

                
                if data['country_Names'][i] == 'Chile':
                    all_page_dataChile = []
                    try:
                        # Find all table rows
                        rows = driver.find_elements(By.XPATH, f'//table[@id="{data["class_names"][i]}"]/tbody/tr')
                        if not rows:
                            print("No rows found on this page.")
                            return all_page_dataChile

                        # Process each row
                        for row in rows:
                            columns = row.find_elements(By.TAG_NAME, "td")
                            row_data = [col.text.strip() for col in columns]

                            # Add 'master_Ids' and 'source_datas' to the row
                            master_id = data['master_Ids'][i] if 'master_Ids' in data else "No master ID"
                            source_data = data['source_datas'][i] if 'source_datas' in data else "No source data"

                            # Combine all data into a dictionary
                            all_page_dataChile.append({
                                'data': row_data,
                                'master_Ids': master_id,
                                'source_datas': source_data
                            })
                    except Exception as e:
                        print(f"Error while scraping table: {e}")
                    return all_page_dataChile

                if data['country_Names'][i] == 'Bulgaria':
                    # Search for the word 'Azithromycin' in the table
                    search_word = data['input_values'][i]
                    matching_rows = []  # To store rows containing the specific word

                    try:
                    # Locate all table rows
                        rows = driver.find_elements(By.XPATH, '//table/tbody/tr')
                        if not rows:
                            print("No rows found in the table.")
                        else:
                            # Iterate through each row
                            for row in rows:
                                # Get all <td> elements in the row
                                columns = row.find_elements(By.TAG_NAME, "td")
                                row_data = [col.text.strip() for col in columns]  # Extract text from each column

                                # Check if the specific word is in the row data
                                if any(search_word in cell for cell in row_data):

                                    # Add 'master_Ids' and 'source_datas' to the row
                                    master_id = data['master_Ids'][i] if 'master_Ids' in data else "No master ID"
                                    source_data = data['source_datas'][i] if 'source_datas' in data else "No source data"

                                    # Combine all data into a dictionary
                                    matching_rows.append({
                                        'data': row_data,
                                        'master_Ids': master_id,
                                        'source_datas': source_data
                                    })
                                    
                    except Exception as e:
                        print(f"Error while scraping the table: {e}")
                    return  matching_rows

                if data['country_Names'][i] == 'France':
                    all_page_dataFrance = []
                    try:
                        if "No results found." not in driver.page_source:
                            # Find all rows containing the medication data
                            rows = driver.find_elements(By.XPATH, "//tr[td//a[@class='standart']]")
                            
                            # Loop through each row and extract data
                            for row in rows:
                                # Find the medication name link in each row
                                medication_link = row.find_element(By.XPATH, ".//a[@class='standart']")
                                medication_name = medication_link.text.strip()

                                # Optionally, you can also check if there is additional information in the row
                                # For example, availability or status info
                                status_cells = row.find_elements(By.XPATH, ".//td[contains(@class, 'ResultRow')]")
                                status_info = [status.text.strip() for status in status_cells]

                                # Store the data for this row
                                row_data = [medication_name] + status_info

                                # Add 'master_Ids' and 'source_datas' to the row
                                master_id = data['master_Ids'][i] if 'master_Ids' in data else "No master ID"
                                source_data = data['source_datas'][i] if 'source_datas' in data else "No source data"

                                # Combine all data into a dictionary
                                all_page_dataFrance.append({
                                    'data': row_data,
                                    'master_Ids': master_id,
                                    'source_datas': source_data
                                })
                    except Exception as e:
                        print(f"Error while scraping the table: {e}")
                    
                    return all_page_dataFrance

                if data['country_Names'][i] == 'Malta':
                    page_dataMalta = []  # Initialize an empty list to store data

                    # Locate all result items on the new page
                    result_items = driver.find_elements(By.CLASS_NAME, "result-item")

                    # Debug: Check if any result items are found
                    print(f"Found {len(result_items)} result items.")

                    # Iterate over each result item
                    for item in result_items:
                        # Temporary list to store data values
                        data_values = []

                        # Extract the title
                        try:
                            title = item.find_element(By.CSS_SELECTOR, ".result-title .title-md h4").text.strip()
                            data_values.append(title)  # Add title to data_values list
                        except Exception as e:
                            print(f"Error extracting title: {e}")
                            continue

                        # Extract various specifications
                        specs = item.find_elements(By.CSS_SELECTOR, ".result-specs")
                        for spec in specs:
                            try:
                                # Only extract the content (assuming label is not required in the list structure)
                                content = spec.find_element(By.CSS_SELECTOR, ".result-content table-item span").text.strip()
                                data_values.append(content)  # Add content to data_values list
                            except Exception as e:
                                print(f"Error extracting specs: {e}")
                                continue  # Skip if any content is not found

                        # Add 'master_Ids' and 'source_datas' to the row
                        try:
                            master_id = data['master_Ids'][i] if 'master_Ids' in data else "No master ID"
                            source_data = data['source_datas'][i] if 'source_datas' in data else "No source data"

                            # Combine all data into a dictionary
                            page_dataMalta.append({
                                'data': data_values,
                                'master_Ids': master_id,
                                'source_datas': source_data
                            })
                        except Exception as e:
                            print(f"Error appending data: {e}")
                            continue

                    # Debug: Check the final data before returning
                    print(f"Collected data: {page_dataMalta}")
                    return page_dataMalta

                if data['country_Names'][i] == 'Turkey':
                    try:
                        # Wait for the first <div> with the specified class to be visible
                        first_div = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CLASS_NAME, "cell.text-center"))  # Adjust the class selector if necessary
                        )
                        
                        # Find the <a> element inside the first <div>
                        xlsx_button = first_div.find_element(By.TAG_NAME, "a")
                        
                        # Click the button
                        xlsx_button.click()
                        print("Clicked the XLSX button successfully.")
                    except NoSuchElementException:
                        print("The specified <div> or <a> button was not found.")
                    except TimeoutException:
                        print("Timeout while waiting for the <div> or <a> button.")
                    except Exception as e:
                        print(f"An error occurred: {e}")


                    # Pagination
            def has_next_page():
                try:
                    next_button = driver.find_element(By.CLASS_NAME, 'ui-paginator-next')
                    return 'disabled' not in next_button.get_attribute('class')
                except NoSuchElementException:
                    return False

            def go_to_next_page():
                try:
                    next_button = driver.find_element(By.CLASS_NAME, 'ui-paginator-next')
                    driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, 'ui-paginator-next')))
                    next_button.click()
                    time.sleep(8)
                except ElementClickInterceptedException:
                    print("Next button click intercepted. Retrying...")
                    try:
                        overlay = driver.find_element(By.CSS_SELECTOR, ".overlay-selector")
                        driver.execute_script("arguments[0].style.visibility = 'hidden';", overlay)
                        next_button.click()
                        time.sleep(8)
                    except NoSuchElementException:
                        print("No overlay found, retrying click.")
                        driver.execute_script("arguments[0].click();", next_button)
                except (NoSuchElementException, ElementNotInteractableException) as e:
                    print(f"Failed to go to next page. Error: {e}")

            # Scrape all pages
            page_number = 1
            while True:
                print(f"Scraping page {page_number}...")
                page_data = scrape_table()
                all_scraped_data.extend(page_data)

                if has_next_page():
                    go_to_next_page()
                    page_number += 1
                else:
                    break

    except Exception as e:
            print(f"Error scraping {base_url}: {e}")

    finally:
        driver.quit()

    return jsonify({"scraped_data": all_scraped_data})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
