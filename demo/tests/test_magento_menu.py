import openpyxl
import logging
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

URL = "https://magento.softwaretestingboard.com/"

# Configure logging
logging.basicConfig(level=logging.INFO)

def test_mega_menu(driver, wait, actions):
    driver.get(URL)
    logging.info("Navigated to the URL")

    #  Hover on 'Gear' option from Header
    gear_menu = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@id='ui-id-6']")))
    actions.move_to_element(gear_menu).perform()
    logging.info("Hovered over 'Gear' menu")

    #  Click on 'Bags' from Gear menu
    try:
        bags_menu = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@id='ui-id-25']")))
        bags_menu.click()
        logging.info("Clicked on 'Bags' menu")
    except TimeoutException:
        logging.error("Failed to find 'Bags' menu", exc_info=True)
        print(driver.page_source)  # Print the page source for debugging
        raise

def test_plp(driver, wait):
    #  Create an Excel sheet for product details
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Product Details"
    ws.append(["Product Name", "Product Price"])

    #  Print all product names and prices on the PLP to an Excel file
    products = wait.until(EC.presence_of_all_elements_located((By.XPATH, "//li[@class='item product product-item']")))

    for product in products:
        product_name = product.find_element(By.CSS_SELECTOR, ".product-item-link").text
        product_price = product.find_element(By.CSS_SELECTOR, ".price").text
        ws.append([product_name, product_price])
        logging.info(f"Added product '{product_name}' with price '{product_price}' to Excel")

    # Save the workbook
    wb.save("Product_Details.xlsx")
    logging.info("Saved product details to Excel")

    #  Click on any product to see product details
    products[0].find_element(By.CSS_SELECTOR, ".product-item-link").click()
    logging.info("Clicked on a product to view details")

def test_pdp(driver, wait):
    # Step 5: Navigate to PDP and check product details page
    pdp_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".product.media")))
    assert pdp_element.is_displayed(), "Product Details Page is not displayed."
    logging.info("Product Details Page is displayed")

    # Change the quantity of the product
    quantity_field = driver.find_element(By.ID, "qty")
    quantity_field.clear()
    quantity_field.send_keys("2")
    logging.info("Changed product quantity to 2")

    # Click on 'Add to Cart' button
    add_to_cart_button = driver.find_element(By.ID, "product-addtocart-button")
    add_to_cart_button.click()
    logging.info("Clicked 'Add to Cart' button")

    # Check the cart count is updated
    cart_count = wait.until(EC.text_to_be_present_in_element((By.CSS_SELECTOR, ".counter.qty"), "2"))
    assert cart_count, "Cart count is not updated."
    logging.info("Cart count is updated to 2")

def test_cart_page(driver, wait):
    # Open cart page through cart icon
    cart_icon = driver.find_element(By.CSS_SELECTOR, ".showcart")
    cart_icon.click()
    logging.info("Clicked on cart icon")

    view_edit_cart_button = wait.until(EC.visibility_of_element_located((By.XPATH, "//a[@class='action viewcart']")))
    view_edit_cart_button.click()
    logging.info("Clicked 'View and Edit Cart' button")

    #  Verify product is displayed and details are correct in cart page
    cart_product_name = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".product-item-name a"))).text
    cart_product_quantity = driver.find_element(By.CSS_SELECTOR, ".cart-item-qty input").get_attribute("value")

    assert cart_product_name, "Product name in cart does not match."
    assert cart_product_quantity == "2", "Product quantity in cart is not correct."
    logging.info("Verified product name and quantity in cart")

    # Change the quantity of the product in cart page
    cart_quantity_field = driver.find_element(By.CSS_SELECTOR, ".cart-item-qty input")
    cart_quantity_field.clear()
    cart_quantity_field.send_keys("3")

    update_cart_button = driver.find_element(By.NAME, "update_cart_action")
    update_cart_button.click()
    logging.info("Changed product quantity to 3 and clicked 'Update Cart' button")

    # Proceed to Checkout
    checkout_button = wait.until(EC.visibility_of_element_located((By.XPATH, "//button[@data-role='proceed-to-checkout']")))
    checkout_button.click()
    logging.info("Clicked 'Proceed to Checkout' button")

def test_checkout(driver, wait):
    #  Fill the checkout details (Assuming guest checkout)
    try:
        first_name = wait.until(EC.visibility_of_element_located((By.NAME, "firstname")))
        first_name.send_keys("Test")
        logging.info("Entered first name")

        last_name = driver.find_element(By.NAME, "lastname")
        last_name.send_keys("User")
        logging.info("Entered last name")

        email = driver.find_element(By.NAME, "email")
        email.send_keys("test@example.com")
        logging.info("Entered email")

        street_address = driver.find_element(By.NAME, "street[0]")
        street_address.send_keys("123 Test St")
        logging.info("Entered street address")

        city = driver.find_element(By.NAME, "city")
        city.send_keys("Testville")
        logging.info("Entered city")

        state = driver.find_element(By.NAME, "region_id")
        state.send_keys("California")
        logging.info("Entered state")

        zip_code = driver.find_element(By.NAME, "postcode")
        zip_code.send_keys("12345")
        logging.info("Entered zip code")

        phone = driver.find_element(By.NAME, "telephone")
        phone.send_keys("5555555555")
        logging.info("Entered phone number")

        # Click on 'Next' button to proceed to payment
        next_button = driver.find_element(By.XPATH, "//button[@class='button action continue primary']")
        next_button.click()
        logging.info("Clicked 'Next' button")

        # Place the order after filling all the details
        place_order_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@class='action primary checkout']")))
        place_order_button.click()
        logging.info("Clicked 'Place Order' button")

        # Check if the order is placed successfully
        order_success_message = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".checkout-success")))
        assert order_success_message.is_displayed(), "Order was not placed successfully."
        logging.info("Order placed successfully")
    except TimeoutException:
        # Take a screenshot for debugging
        driver.save_screenshot("checkout_error.png")
        logging.error("Timeout occurred during checkout", exc_info=True)
        print(driver.page_source)  # Print the page source for debugging
        raise
